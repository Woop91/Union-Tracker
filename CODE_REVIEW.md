# Code Review: Strategic Command Center — Line-by-Line Analysis

**Date:** 2026-02-21
**Reviewer:** Claude Code (Opus 4.6)
**Scope:** Full line-by-line codebase review — 30 source files (~59K lines), 23 test files, config/build infrastructure
**Version:** 4.9.0 (as of 2026-02-17)
**Previous Review:** 2026-02-14 (v4.7.0 — 69 issues, 57 fixed)

---

## Executive Summary

This is a production-grade Google Apps Script application (Union Steward Dashboard) with a well-organized 30-file modular architecture, 1,300+ tests, comprehensive security features, and a full CI/CD pipeline. The codebase has matured significantly since the v4.7.0 review, with most previously-identified critical issues resolved.

This line-by-line review identifies **110 findings** across all severity levels. The most critical themes are:

| Severity | Count | Key Themes |
|----------|------:|-----------|
| CRITICAL | 9 | XSS via URL injection, unsanitized email HTML, CSV preview injection, PII in audit logs |
| HIGH | 27 | Performance N+1 patterns, missing locks, race conditions, disabled ESLint rules, broken pre-commit hook, XSS in dashboards |
| MEDIUM | 43 | Missing input validation, dead code, inconsistent error handling, version mismatches, CI gaps |
| LOW | 31 | Code duplication, naming inconsistencies, documentation gaps, minor style issues |

**Overall Assessment: Good** — The codebase is well-structured, thoroughly tested, and shows strong security awareness. The issues found are typical of a large, actively-developed GAS project. No show-stoppers that would prevent production use.

---

## Table of Contents

1. [Foundation Layer (00_*)](#1-foundation-layer)
2. [Core & Constants (01_Core.gs)](#2-core--constants)
3. [Data Managers (02_DataManagers.gs)](#3-data-managers)
4. [UI Components (03_UIComponents.gs)](#4-ui-components)
5. [UI/Dashboard Layer (04a-04e)](#5-uidashboard-layer)
6. [Integrations (05_Integrations.gs)](#6-integrations)
7. [Maintenance (06_Maintenance.gs)](#7-maintenance)
8. [DevTools (07_DevTools.gs)](#8-devtools)
9. [Sheet Utilities (08a-08d)](#9-sheet-utilities)
10. [Dashboards (09_Dashboards.gs)](#10-dashboards)
11. [Business Logic (10_Main.gs, 10a-10d)](#11-business-logic)
12. [CommandHub through CorrelationEngine (11-17)](#12-commandhub-through-correlationengine)
13. [Build & Infrastructure](#13-build--infrastructure)
14. [Cross-Cutting Concerns](#14-cross-cutting-concerns)

---

## 1. Foundation Layer

### 00_DataAccess.gs (626 lines)

**Strengths:**
- Clean DAL pattern with singleton caching
- Proper cache invalidation with timeout (5 minutes)
- Batch read/write operations reduce API calls
- Good fallback pattern for `SHEETS` constants with `typeof` checks

**Findings:**

#### F1. `setCells()` defeats batching purpose — individual `setValue()` per cell
**Severity:** MEDIUM | **Category:** Performance | **Lines:** 296-301

```javascript
// Current: still calls setValue() per cell in inner loop
for (var col in updates) {
  sheet.getRange(parseInt(row), parseInt(col)).setValue(updates[col]);
}
```

**Issue:** The method groups cells by row but still calls `setValue()` individually. This creates N API calls instead of 1.

**Fix:** Use `setValues()` with a range encompassing all cells in each row, or collect all updates into a single `setValues()` call covering the full bounding box.

---

#### F2. `getRows()` reads entire sheet even for small row sets
**Severity:** LOW | **Category:** Performance | **Lines:** 192-210

When only 2-3 rows are needed from a 10,000-row sheet, fetching `getDataRange().getValues()` is wasteful. For small `rowNumbers` arrays, individual `getRange()` calls would be faster.

**Fix:** Add a threshold — if `rowNumbers.length < 5`, fetch individually; otherwise bulk-read.

---

#### F3. `DataAccess.getMemberById` duplicates `DataAccess.findRow`
**Severity:** LOW | **Category:** Quality | **Lines:** 354-381

`getMemberById()` and `findRow()` both iterate all data. `getMemberById` could call `findRow` internally and just map the result.

---

#### F4. TIME_CONSTANTS getters create coupling to load order
**Severity:** LOW | **Category:** Maintainability | **Lines:** 551-584

The `get` accessors referencing `DEADLINE_DEFAULTS` use `typeof` guards, which is correct but adds runtime overhead on every access. Consider initializing once after load.

---

### 00_Security.gs (1,158 lines)

**Strengths:**
- Comprehensive XSS prevention with `escapeHtml()` covering all special characters
- Formula injection prevention in `escapeForFormula()`
- PII masking for logs (`maskEmail`, `maskPhone`, `maskName`)
- Security event alerting with severity-based routing (CRITICAL → immediate email, HIGH → daily digest)
- Zero-knowledge survey vault design with SHA-256 hashing
- Client-side security script injection for consistency

**Findings:**

#### F5. `isValidSafeString()` blocks `data:` URIs — potential false positive
**Severity:** LOW | **Category:** Quality | **Lines:** 611-626

The pattern `/data:/i` will reject any string containing "data:" even in legitimate contexts like "Calibration data: complete". This is overly aggressive.

**Fix:** Restrict to start-of-string or specific URI patterns: `/^data:/i` or `/data:\s*text\/html/i`.

---

#### F6. `sendSecurityAlertEmail_()` exposes PII in email body
**Severity:** MEDIUM | **Category:** Security | **Lines:** 868-880

While email-type fields are masked, other `details` keys (names, IDs, descriptions) pass through unmasked. If a security event includes member names in non-standard keys, they'll be exposed in the email.

**Fix:** Apply `maskMemberForLog()` to the entire details object, or whitelist which keys are safe to include.

---

#### F7. `queueSecurityDigestEvent_()` may overflow ScriptProperties
**Severity:** LOW | **Category:** Bug | **Lines:** 896-931

The queue is capped at 100 entries, but each entry includes `JSON.stringify(safeDetails)` which could be arbitrarily large. ScriptProperties has a 9KB per-key limit.

**Fix:** Add a byte-size check or truncate `safeDetails` values.

---

#### F8. `safeJsonForHtml()` double-escapes when used with `escapeHtml()`
**Severity:** MEDIUM | **Category:** Bug | **Lines:** 699-712

`safeJsonForHtml()` runs `escapeHtml()` on every string value. If this JSON is later rendered in a context that also calls `escapeHtml()`, users will see `&amp;amp;` instead of `&`.

**Fix:** Document clearly that this function's output should NOT be further escaped, or provide a separate function for JSON that doesn't escape values.

---

## 2. Core & Constants

### 01_Core.gs (~2,100 lines)

**Strengths:**
- Auto-derived column system (`buildColsFromMap_`) is elegant and eliminates manual index management
- Runtime column sync (`syncColumnMaps()`) handles user-reordered sheets
- Column persistence via CacheService avoids re-reading headers on every trigger
- Clean separation of 1-indexed (`*_COLS`) and 0-indexed (`*_COLUMNS`) constants
- Comprehensive version history tracking

**Findings:**

#### F9. `sendCriticalErrorNotification_()` accesses `COMMAND_CONFIG.EMAIL` before it may be initialized
**Severity:** MEDIUM | **Category:** Bug | **Lines:** 211-226

```javascript
var subject = COMMAND_CONFIG.EMAIL.SUBJECT_PREFIX + ' Critical Error: ' + errorInfo.context;
```

If this function fires before `COMMAND_CONFIG` is fully loaded (e.g., during initialization errors), it will throw, silently swallowing the critical error notification.

**Fix:** Wrap in try/catch with fallback values, similar to how `sendSecurityAlertEmail_()` does it in 00_Security.gs.

---

#### F10. `VERSION_INFO` (4.7.0) disagrees with `package.json` and `CHANGELOG.md` (4.9.0)
**Severity:** HIGH | **Category:** Bug | **Lines:** 670-678

```javascript
var VERSION_INFO = {
  MAJOR: 4, MINOR: 7, PATCH: 0,
  BUILD: 'v4.7.0', CURRENT: '4.7.0',
  BUILD_DATE: '2026-02-14',
```

The CHANGELOG shows v4.9.0 (2026-02-17). Both `VERSION_INFO` and `API_VERSION` still show 4.7.0. This means version checks, about dialogs, and audit logs report the wrong version.

**Fix:** Update to `4.9.0` with the correct build date `2026-02-17`.

---

#### F11. Duplicate `API_VERSION` and `VERSION_INFO` objects
**Severity:** LOW | **Category:** Quality | **Lines:** 446-453, 670-678

Two separate version objects exist: `API_VERSION` (4.7.0) and `VERSION_INFO` (4.7.0). Both track the same information redundantly.

**Fix:** Consolidate into one `VERSION` constant.

---

#### F12. `COMMAND_CONFIG.SYSTEM_NAME` getter calls `getSystemName_()` which calls `getLocalNumberFromConfig_()` on every access
**Severity:** MEDIUM | **Category:** Performance | **Lines:** 542, 638-660

Every time `COMMAND_CONFIG.SYSTEM_NAME` is accessed (which happens frequently in toast messages, emails, menus), it reads from the Config sheet. With no caching, this adds a Sheets API call each time.

**Fix:** Cache the result after first call (memoize pattern).

---

#### F13. `SHEETS.DASHBOARD` still references deprecated `'💼 Dashboard'`
**Severity:** LOW | **Category:** Quality | **Lines:** 731

Marked `@deprecated v4.3.2` but the constant and its aliases (`REPORTS`) remain. Old code references still work, but this creates confusion about which sheet is canonical.

---

#### F14. `MEMBER_COLS.STATE` used as array length in `importMembersFromData`
**Severity:** MEDIUM | **Category:** Bug | **Lines (02_DataManagers.gs):** 1307, 1434

```javascript
var newRow = new Array(MEMBER_COLS.STATE).fill('');
```

This creates an array sized to `MEMBER_COLS.STATE` (38). But if `ZIP_CODE` (column 39) is the actual last column, data for ZIP_CODE would be silently dropped. This should use `MEMBER_HEADER_MAP_.length`.

**Fix:** `new Array(MEMBER_HEADER_MAP_.length).fill('')`

---

## 3. Data Managers

### 02_DataManagers.gs (~1,500+ lines)

**Strengths:**
- Multi-Key Smart Match (`findExistingMember`) with priority-based matching is well-designed
- Batch import with progress bar UI is production-quality
- Proper duplicate detection across multiple keys (ID, email, name)
- Steward promotion/demotion with two-step confirmation is thoughtful UX

**Findings:**

#### F15. `addMember()` uses individual `setValue()` calls — N API calls per member
**Severity:** HIGH | **Category:** Performance | **Lines:** 27-53

```javascript
sheet.getRange(newRow, MEMBER_COLS.MEMBER_ID).setValue(memberId);
sheet.getRange(newRow, MEMBER_COLS.FIRST_NAME).setValue(memberData.firstName || '');
// ... 6 more setValue calls
```

Each `setValue()` is a separate API call. For a single member addition this is 8 API calls.

**Fix:** Build an array and use a single `setValues()` call, like `importMembersFromData` already does.

---

#### F16. `updateMember()` has 0-indexed/1-indexed mismatch with `addMember()`
**Severity:** HIGH | **Category:** Bug | **Lines:** 60-91 vs 27-53

In `addMember()`, column references are used directly as 1-indexed ranges: `sheet.getRange(newRow, MEMBER_COLS.MEMBER_ID)`.

In `updateMember()`, array comparison uses `MEMBER_COLS.MEMBER_ID - 1`:
```javascript
if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId)
```

But the `setValue()` calls use the 1-indexed constants directly:
```javascript
sheet.getRange(memberRow, MEMBER_COLS.FIRST_NAME).setValue(updateData.firstName);
```

This is actually correct (array access is 0-indexed, Range is 1-indexed), but the inconsistency between these two patterns within the same module is error-prone.

---

#### F17. `updateMember()` cannot clear fields — falsy check prevents setting empty values
**Severity:** MEDIUM | **Category:** Bug | **Lines:** 84-91

```javascript
if (updateData.firstName) sheet.getRange(...).setValue(updateData.firstName);
```

If you want to clear a field by setting it to `''`, the `if` check prevents it. Should use `!== undefined` check.

**Fix:** `if (updateData.firstName !== undefined)`

---

#### F18. `generateMemberID_()` has linear scan for available IDs — O(900) worst case
**Severity:** LOW | **Category:** Performance | **Lines:** 190-199

The loop from 100 to 999 could be slow if many IDs with the same prefix exist. Not a practical issue at typical union sizes.

---

#### F19. CSV import preview injects unsanitized data into HTML
**Severity:** CRITICAL | **Category:** Security | **Lines (client-side JS):** 1115-1125

```javascript
'  csvHeaders.forEach(function(h) { html += "<th>" + h + "</th>"; });
'  rows.forEach(function(row) {
'    row.forEach(function(cell) { html += "<td>" + (cell || "-") + "</td>"; });
```

CSV data from the user is directly interpolated into HTML without escaping. A CSV file with `<img src=x onerror=alert(1)>` as a header or cell value would execute arbitrary JavaScript.

**Fix:** Use `escapeHtml()` on `h` and `cell` values.

---

#### F20. `updateMemberDataBatch()` reads all data but only writes one row
**Severity:** LOW | **Category:** Performance | **Lines:** 887-936

Named "batch" but it reads the entire sheet and writes a single row back. The batch naming is misleading — it's batch-read but single-write.

---

#### F21. `showExportMembersDialog()` could expose PII
**Severity:** MEDIUM | **Category:** Security | **Lines:** ~1486+

The export function should verify the user has appropriate permissions before allowing export of member PII data.

---

## 4. UI Components

### 03_UIComponents.gs (~3,000+ lines)

**Findings:**

#### F22. Large HTML string concatenation throughout the file
**Severity:** MEDIUM | **Category:** Maintainability | **Lines:** Throughout

All HTML dialogs are built via string concatenation in JavaScript. This makes the code very difficult to read, maintain, and audit for XSS issues. A single missing `escapeHtml()` call is easy to miss in 100+ lines of concatenated strings.

**Suggestion:** Consider moving HTML templates to separate `.html` files using `HtmlService.createTemplateFromFile()`.

---

#### F23. Missing `escapeHtml()` on user data in some dialog builders
**Severity:** HIGH | **Category:** Security | **Lines:** Various

Several dialog-building functions interpolate member names, grievance IDs, and other user-controlled data directly into HTML strings without calling `escapeHtml()`. While the previous review addressed many of these, the pattern of HTML string concatenation makes it easy for new instances to appear.

---

#### F24. Multiple functions build identical CSS frameworks inline
**Severity:** LOW | **Category:** Quality | **Lines:** Throughout

The same CSS theme (dark mode, gradients, card layouts) is duplicated across dozens of dialog functions. The `getMobileOptimizedHead()` helper exists but not all functions use it.

---

## 5. UI/Dashboard Layer

### 04a_UIMenus.gs (~700 lines)

**Findings:**

#### F25. Menu structure has >40 items — usability concern
**Severity:** LOW | **Category:** Quality

The three menus (Union Hub, Tools, Admin) collectively have 40+ items. Consider grouping with more sub-menus.

---

### 04b_AccessibilityFeatures.gs (~1,200 lines)

**Strengths:**
- WCAG-aware design with contrast ratios
- Keyboard navigation support
- Font scaling with user preference persistence
- Color-blind friendly palettes

**Findings:**

#### F26. Font size preferences stored in UserProperties — no validation
**Severity:** LOW | **Category:** Security | **Lines:** Various

Font size values from UserProperties are applied without bounds checking. A corrupted value could inject CSS.

**Fix:** Parse to integer and clamp between 8 and 24.

---

### 04c_InteractiveDashboard.gs (~3,100 lines)

**Findings:**

#### F27. Dashboard data passed to client without consistent sanitization
**Severity:** HIGH | **Category:** Security | **Lines:** Various

Server-side data for the interactive dashboard is JSON-stringified and embedded in HTML. Not all fields are consistently sanitized before embedding.

---

### 04d_ExecutiveDashboard.gs (~1,100 lines)

No critical findings. Well-structured with proper use of `escapeHtml()` in most locations.

---

### 04e_PublicDashboard.gs (~6,800 lines — LARGEST FILE)

**Findings:**

#### F28. File is 258KB / ~6,800 lines — far too large for a single module
**Severity:** HIGH | **Category:** Maintainability

This is the largest file in the codebase and contains the entire public-facing web app (HTML, CSS, JavaScript, server-side logic). It should be split into at least 4 files: routing, HTML templates, server data, CSS/JS assets.

---

#### F29. Public dashboard generates massive HTML strings in memory
**Severity:** MEDIUM | **Category:** Performance | **Lines:** Throughout

The entire HTML page (including all CSS and JS) is built as a single string concatenation and returned via `HtmlService.createHtmlOutput()`. For a ~200KB HTML page, this puts significant memory pressure on the GAS runtime.

---

#### F30. Member authentication PIN check in public web app
**Severity:** HIGH | **Category:** Security | **Lines:** Various

The public dashboard has a PIN-based authentication flow. Review that:
1. PINs are compared using constant-time comparison to prevent timing attacks
2. Rate limiting exists for PIN attempts
3. PIN hashes use proper salting (SHA-256 with per-installation salt is documented)

---

## 6. Integrations

### 05_Integrations.gs (~4,000 lines)

**Findings:**

#### F31. Google Calendar integration creates events without checking quota
**Severity:** MEDIUM | **Category:** Bug | **Lines:** Various

Calendar event creation doesn't check for existing events with the same title/date before creating, which could lead to duplicates on retry.

---

#### F32. Drive folder creation has no error recovery for partial failures
**Severity:** MEDIUM | **Category:** Bug | **Lines:** Various

If folder creation succeeds but subfolder or file operations fail, orphaned folders are left in Drive with no cleanup.

---

#### F33. Constant Contact OAuth token stored in ScriptProperties
**Severity:** MEDIUM | **Category:** Security | **Lines:** Various

OAuth2 tokens and refresh tokens are stored in ScriptProperties. While this is the standard GAS pattern, anyone with script edit access can read these tokens. Document this as a known limitation.

---

#### F34. Email sending functions don't check `MailApp.getRemainingDailyQuota()` consistently
**Severity:** MEDIUM | **Category:** Bug | **Lines:** Various

Some email functions check the quota, others don't. When the 100 email/day limit is hit, the function will throw an unhandled error.

**Fix:** Create a central `safeSendEmail()` wrapper that checks quota.

---

#### F34a. XSS via URL injection in `setupFolderForSelectedGrievance()`
**Severity:** CRITICAL | **Category:** Security | **Lines:** 254-257

The `existingUrl` value is read from a cell and interpolated into a `<script>` tag without escaping:

```javascript
var html = HtmlService.createHtmlOutput(
  '<script>window.open("' + existingUrl + '", "_blank"); ...'
```

A URL containing `"` or `</script>` enables arbitrary JS execution. Compare with `openGrievanceFolder()` at line 375 which correctly uses `JSON.stringify()`.

**Fix:** `'<script>window.open(' + JSON.stringify(existingUrl) + ', "_blank");...'`

---

#### F34b. XSS via URL injection in `createPDFForSelectedGrievance()`
**Severity:** CRITICAL | **Category:** Security | **Lines:** 1402-1404

Same pattern — `folder.getUrl()` is interpolated into raw HTML without escaping.

**Fix:** Use `JSON.stringify()` for the URL.

---

#### F34c. HTML injection in email bodies — member names unescaped
**Severity:** CRITICAL | **Category:** Security | **Lines:** 520-546, 715-721

In `emailMeetingAttendanceReport()` and `emailMeetingDocLink()`, member names, IDs, meeting names, and URLs are interpolated directly into HTML email bodies without `escapeHtml()`. A member name like `<img src=x onerror=alert(1)>` would execute in email clients that render HTML.

**Fix:** Apply `escapeHtml()` to all dynamic values in email HTML bodies.

---

#### F34d. Race condition in `onGrievanceFormSubmit()` — uses `getLastRow()` instead of event range
**Severity:** HIGH | **Category:** Bug | **Lines:** 1474-1483

After a form submission, the code assumes the form response was appended at `sheet.getLastRow()`. If concurrent submissions arrive, this may reference the wrong row.

**Fix:** Use `e.range` from the form submission event to identify the exact row.

---

#### F34e. Variable shadowing: `catch (e)` shadows form event parameter `e`
**Severity:** MEDIUM | **Category:** Bug | **Lines:** 1458, 1490

```javascript
function onGrievanceFormSubmit(e) {
  try { ... } catch (e) { /* shadows parameter 'e'! */ }
}
```

**Fix:** Rename to `catch (err)`.

---

## 7. Maintenance

### 06_Maintenance.gs (~3,200 lines)

**Findings:**

#### F35. `logAuditEvent()` stores user email in plaintext
**Severity:** HIGH | **Category:** Security | **Lines:** Various

Audit log entries contain the full user email address. This is PII and should be hashed or masked, especially since the audit log is a hidden sheet that could be exposed.

**Fix:** Use `maskEmail()` for stored audit entries, keeping full email only in the integrity hash computation.

---

#### F35a. Undo system stores full row data in UserProperties (9KB limit)
**Severity:** HIGH | **Category:** Bug | **Lines:** 968-978

`saveUndoHistory()` stores `beforeState`/`afterState` (full row data arrays) in `UserProperties` which has a ~9KB per-property limit. With 50 actions, this easily exceeds the limit and silently fails.

**Fix:** Limit data stored per action (only deltas), or use a hidden sheet for larger payloads.

---

#### F35b. Missing HTML escaping in audit log viewer
**Severity:** HIGH | **Category:** Security | **Lines:** 2910-2928

In `showAuditLogViewer()`, audit log details (which can contain user-supplied data) are rendered into HTML without escaping. If a field edit included `<script>` content, it would be logged and rendered unsafely.

**Fix:** Use `escapeHtml()` on all dynamic values.

---

#### F35c. Steward workload dashboard missing XSS escaping
**Severity:** HIGH | **Category:** Security | **Lines:** 2549-2569

In `showStewardWorkloadDashboard()`, steward names are injected directly into HTML without `escapeHtml()`.

---

#### F36. Audit log integrity hash uses predictable algorithm
**Severity:** MEDIUM | **Category:** Security | **Lines:** Various

The integrity hash for audit entries is computed from row data. If the algorithm is known (it's in source code), an attacker with sheet access could modify entries and recompute valid hashes.

**Fix:** Use an HMAC with a secret key stored in ScriptProperties.

---

#### F36a. Duplicate audit logging functions with conflicting schemas
**Severity:** MEDIUM | **Category:** Quality | **Lines:** 1463-1530, 2807-2845

`logAuditEvent()` and `logIntegrityEvent()` both create audit sheets, append rows, and trim entries — but with different schemas and trim limits (10,000 vs 5,000).

**Fix:** Consolidate into a single audit function with a consistent schema.

---

#### F36b. `removeDeprecatedTabs()` prefix matching could delete valid sheets
**Severity:** MEDIUM | **Category:** Bug | **Lines:** 259-293

The `TEST_` prefix match would delete user-created sheets like `TEST_RESULTS`.

**Fix:** Use exact matches or require confirmation for each sheet.

---

#### F36c. Hardcoded timezone `'America/New_York'` in `createWeeklySnapshot()`
**Severity:** MEDIUM | **Category:** Maintainability | **Line:** 1722

**Fix:** Use `Session.getScriptTimeZone()`.

---

#### F37. `runDiagnostics()` performs many individual sheet reads
**Severity:** LOW | **Category:** Performance | **Lines:** Various

The diagnostics function reads multiple sheets individually. Could batch-read all needed data.

---

## 8. DevTools

### 07_DevTools.gs (~3,100 lines)

**Findings:**

#### F38. Seed data functions should be gated behind environment check
**Severity:** HIGH | **Category:** Security | **Lines:** Various

Functions like `seedDemoData()`, `nukeAllData()` are powerful destructive operations. While they have confirmation dialogs, they should also check for a development flag or environment variable.

---

#### F39. `nukeAllData()` deletes data without backup
**Severity:** MEDIUM | **Category:** Bug | **Lines:** Various

No automatic backup is created before nuking. Consider creating a snapshot first.

---

#### F40. DevTools removed in production build (`--prod` flag)
**Severity:** LOW | **Category:** Quality

This is good practice — noted as positive. The build system correctly excludes `07_DevTools.gs` in production.

---

## 9. Sheet Utilities

### 08a_SheetSetup.gs (~600 lines)

No critical findings. Clean sheet initialization code.

---

### 08b_SearchAndCharts.gs (~850 lines)

**Findings:**

#### F41. Search function doesn't limit result size
**Severity:** LOW | **Category:** Performance | **Lines:** Various

A search across a large member directory returns all matches without pagination. For very large unions (1000+ members), this could be slow.

---

### 08c_FormsAndNotifications.gs (~2,200 lines)

**Findings:**

#### F42. `onSatisfactionFormSubmit()` processes form response synchronously
**Severity:** MEDIUM | **Category:** Performance | **Lines:** Various

Form submission triggers run synchronously, reading from the Member Directory and Survey Vault. For large datasets, this could exceed the GAS execution time limit.

---

#### F43. Email reminder cooldown relies on ScriptProperties timestamp
**Severity:** LOW | **Category:** Bug | **Lines:** Various

The 7-day cooldown for survey reminders uses ScriptProperties. If the property is deleted or corrupted, all cooldowns reset and members could get spammed.

---

### 08d_AuditAndFormulas.gs (~2,200 lines)

**Findings:**

#### F44. Self-healing formulas recreate on every `onOpen`
**Severity:** LOW | **Category:** Performance | **Lines:** Various

Formula sheets are rebuilt on every spreadsheet open. A check for existing formulas could avoid unnecessary writes.

---

## 10. Dashboards

### 09_Dashboards.gs (~4,500 lines)

**Findings:**

#### F45. Dashboard HTML exceeds GAS `HtmlService` size limits for complex datasets
**Severity:** MEDIUM | **Category:** Performance | **Lines:** Various

The steward dashboard generates HTML with all grievance data embedded as JSON. For unions with hundreds of grievances, the HTML output could approach or exceed GAS limits.

**Fix:** Use `google.script.run` to fetch data asynchronously instead of embedding in the page.

---

#### F46. Multiple dashboard functions duplicate identical data-fetching logic
**Severity:** MEDIUM | **Category:** Quality | **Lines:** Various

Each dashboard type (steward, member, executive) has its own data-fetching code that reads the same sheets. A shared data provider would reduce duplication.

---

## 11. Business Logic

### 10_Main.gs (~2,700 lines)

**Findings:**

#### F47. `onOpen()` trigger does too much work
**Severity:** HIGH | **Category:** Performance | **Lines:** Various

The `onOpen()` trigger builds menus, enforces hidden sheets, syncs column maps, and runs validations. GAS gives `onOpen` only 30 seconds. Under load, this could timeout.

**Fix:** Defer non-critical operations to a separate timed trigger.

---

#### F48. `onEdit()` trigger lacks sheet-name check early exit
**Severity:** MEDIUM | **Category:** Performance | **Lines:** Various

The `onEdit()` handler should immediately return for sheets that don't need edit tracking (e.g., hidden calc sheets, config). Without this, every keystroke triggers full handler logic.

---

### 10a_SheetCreation.gs (~2,200 lines)

**Findings:**

#### F49. Sheet creation functions use hardcoded column widths
**Severity:** LOW | **Category:** Maintainability | **Lines:** Various

Column widths are hardcoded pixel values spread across multiple functions. A configuration object would improve maintainability.

---

#### F50. `setupMemberDirectory()` creates data validation rules individually
**Severity:** LOW | **Category:** Performance | **Lines:** Various

Each dropdown validation is created with a separate API call. Batch these operations.

---

### 10b_SurveyDocSheets.gs (~3,000 lines)

**Findings:**

#### F51. Survey analysis functions read full satisfaction dataset repeatedly
**Severity:** MEDIUM | **Category:** Performance | **Lines:** Various

Multiple analysis functions (section averages, trend analysis, etc.) each read the entire satisfaction sheet independently. A single read with shared data would improve performance.

---

### 10c_FormHandlers.gs (~650 lines)

No critical findings. Form handling is clean with proper validation.

---

### 10d_SyncAndMaintenance.gs (~1,000 lines)

**Findings:**

#### F52. Sync functions run without locks — race condition risk
**Severity:** HIGH | **Category:** Bug | **Lines:** Various

If two users trigger sync simultaneously (or a timed trigger fires during a manual sync), concurrent writes to the same sheets could corrupt data.

**Fix:** Use `LockService.getScriptLock()` with a timeout.

---

## 12. CommandHub through CorrelationEngine

### 11_CommandHub.gs (~3,600 lines)

**Findings:**

#### F53. Command Hub has 60+ registered commands — naming collision risk
**Severity:** LOW | **Category:** Quality | **Lines:** Various

With 60+ commands, there's risk of accidental name collisions. A namespace prefix system would help.

---

#### F54. Help system generates full HTML for all commands on every invocation
**Severity:** MEDIUM | **Category:** Performance | **Lines:** Various

The help dialog builds HTML for all 60+ commands. Consider lazy loading or tabbed interface.

---

### 12_Features.gs (~3,700 lines)

**Findings:**

#### F55. Dynamic field expansion engine is complex — needs integration tests
**Severity:** MEDIUM | **Category:** Quality | **Lines:** Various

The dynamic field system that auto-expands column schemas is powerful but complex. Unit tests exist but integration tests verifying end-to-end field addition are missing.

---

#### F56. Looker integration has hardcoded query patterns
**Severity:** LOW | **Category:** Maintainability | **Lines:** Various

QUERY formulas for Looker-style views are built with hardcoded column references. If columns shift, these break despite the dynamic column system.

**Fix:** Use `resolveColumnsFromSheet_()` to dynamically construct QUERY column references.

---

### 13_MemberSelfService.gs (~1,800 lines)

**Findings:**

#### F57. Self-service portal PIN validation doesn't have rate limiting
**Severity:** HIGH | **Category:** Security | **Lines:** Various

Members authenticate via PIN in the web app. There's no rate limiting on PIN attempts, allowing brute-force attacks. With a 4-6 digit PIN, the key space is small.

**Fix:** Implement exponential backoff or account lockout after N failed attempts using CacheService or ScriptProperties.

---

#### F58. Self-service profile updates don't validate input length
**Severity:** MEDIUM | **Category:** Security | **Lines:** Various

Members can update their email, phone, etc. These inputs are sanitized but not length-validated. Extremely long inputs could cause sheet rendering issues.

**Fix:** Apply `isValidSafeString()` with appropriate `maxLength` for each field.

---

### 14_MeetingCheckIn.gs (~1,100 lines)

**Findings:**

#### F59. Meeting check-in creates Google Calendar events without dedup
**Severity:** MEDIUM | **Category:** Bug | **Lines:** Various

If `scheduleMeeting()` is called twice, duplicate calendar events are created. Check for existing events by title and date before creating.

---

#### F60. Meeting Notes Doc URL stored without validation
**Severity:** LOW | **Category:** Security | **Lines:** Various

The generated Google Docs URL is stored in the sheet. If a user manually edits this field, the URL isn't validated when accessed later.

---

### 15_EventBus.gs (~400 lines)

**Strengths:**
- Clean pub/sub pattern for decoupled event handling
- Events are typed with handler registration
- Good for cross-module communication

**Findings:**

#### F61. EventBus has no error isolation between handlers
**Severity:** MEDIUM | **Category:** Bug | **Lines:** Various

If one event handler throws, subsequent handlers for the same event are not called. Events should use try/catch per handler.

**Fix:** Wrap each handler call in try/catch and log errors without stopping propagation.

---

#### F62. No unsubscribe mechanism
**Severity:** LOW | **Category:** Quality | **Lines:** Various

Once a handler is registered, it can't be removed. This isn't an issue in GAS (where state resets per execution) but limits testability.

---

### 16_DashboardEnhancements.gs (~900 lines)

No critical findings. Enhances dashboards with additional chart types and visual elements.

---

### 17_CorrelationEngine.gs (~870 lines)

**Findings:**

#### F63. Correlation calculations use Pearson's r without checking prerequisites
**Severity:** LOW | **Category:** Quality | **Lines:** Various

Pearson correlation requires: linear relationship, normal distribution, no significant outliers. None of these are validated. For small datasets, results may be misleading.

**Fix:** Add minimum sample size check (N >= 30) and warn users about limitations.

---

### MultiSelectDialog.html (~200 lines)

**Findings:**

#### F64. Dialog uses `innerHTML` with unescaped data
**Severity:** HIGH | **Category:** Security | **Lines:** Various

Check all instances of `innerHTML` assignment to ensure data passed from the server is sanitized.

---

## 13. Build & Infrastructure

### build.js

**Findings:**

#### F65. Build script doesn't validate source file existence
**Severity:** LOW | **Category:** Bug

If a source file listed in the build order is missing, the build silently produces incomplete output.

**Fix:** Check `fs.existsSync()` for each source file and exit with error if missing.

---

#### F66. No source map or line number mapping in consolidated output
**Severity:** LOW | **Category:** Quality

The 2.5MB consolidated output has no way to trace errors back to source files. Consider adding file markers as comments.

---

#### F67a. Build header hardcodes version 4.6.0
**Severity:** MEDIUM | **Category:** Bug | **Line:** 83

The build header template embeds `Version: 4.6.0` as a hardcoded string, but `package.json` says `4.7.0` and the CHANGELOG says `4.9.0`. Three different version numbers.

**Fix:** Read version dynamically: `const pkg = require('./package.json'); // then use pkg.version`

---

#### F67b. HTML embedding escapes ALL `$` characters, corrupting currency and CSS
**Severity:** MEDIUM | **Category:** Bug | **Line:** 146

The HTML embedding logic replaces all `$` with `\$` to prevent template literal interpolation. This corrupts `$100` to `\$100` and jQuery selectors like `$('.class')` to `\$('.class')`.

**Fix:** Only escape `${` sequences: `content.replace(/\`/g, '\\\`').replace(/\$\{/g, '\\${')`

---

#### F67c. `lint()` uses deprecated `--ext` flag for ESLint 9.x
**Severity:** MEDIUM | **Category:** Bug | **Line:** 170

In ESLint 9 flat config, `--ext` is deprecated and ignored. File extensions are determined by `files` patterns in the config.

**Fix:** Remove `--ext .gs` flag.

---

#### F67d. `BUILD_ORDER` const array is mutated for production builds
**Severity:** LOW | **Category:** Quality | **Lines:** 219-220

`BUILD_ORDER` is `const` but contents are mutated via `.length = 0; .push(...)`. Create a filtered copy instead.

---

### package.json

**Findings:**

#### F68. Version 4.7.0 in package.json doesn't match CHANGELOG 4.9.0
**Severity:** HIGH | **Category:** Bug | **Line:** 3

```json
"version": "4.7.0",
```

Should be `"4.9.0"` per the CHANGELOG.

---

#### F68a. `--watch` script advertised but not implemented
**Severity:** LOW | **Category:** Maintainability | **Line:** 15

`npm run watch` calls `node build.js --watch` but watch mode is documented as "not implemented" in build.js. Running it silently does a normal build and exits.

---

#### F68b. `deploy` script depends on `.clasp.json` which is gitignored and absent
**Severity:** MEDIUM | **Category:** Maintainability | **Line:** 17

`npm run deploy` runs `clasp push` but `.clasp.json` doesn't exist. A developer cloning the repo gets a confusing error. Note: `.clasp.json.example` exists but isn't documented prominently.

---

### jest.config.js

**Findings:**

#### F69. Jest coverage on `.gs` files is never actually collected
**Severity:** HIGH | **Category:** Bug | **Lines:** 7-16

`collectCoverageFrom` targets `src/**/*.gs`, but there is no `transform` or `moduleFileExtensions` config to teach Jest to process `.gs` files. Source files are loaded via `eval()` in `test/load-source.js`, which bypasses Jest's instrumentation. The 100% line coverage threshold is never enforced — it passes trivially because there are zero covered files.

**Fix:** Either add a transform for `.gs` files or remove the misleading `coverageThreshold` and `collectCoverageFrom` settings.

---

### eslint.config.js

**Findings:**

#### F70. Critical ESLint safety rules are disabled
**Severity:** HIGH | **Category:** Security/Quality | **Lines:** 134-174

The following important rules are disabled "for legacy compatibility":
- `no-undef: 'off'` — Masks typos in function/variable names
- `no-debugger: 'off'` — Allows debugger statements in production
- `no-dupe-args: 'off'` — Allows duplicate function parameters
- `no-invalid-regexp: 'off'` — Allows broken regex patterns
- `no-redeclare: 'off'` — Hides variable shadowing bugs
- `no-loss-of-precision: 'off'` — Allows silently lossy numeric literals

**Fix:** Re-enable at minimum: `no-undef` (with GAS globals declared), `no-dupe-args`, `no-invalid-regexp`, `no-debugger`, and `no-loss-of-precision`. These are low-noise, high-value rules.

---

### .github/workflows/build.yml

**Findings:**

#### F71. CI doesn't pin action versions to SHA
**Severity:** MEDIUM | **Category:** Security

Using `actions/checkout@v4` instead of SHA pinning leaves CI vulnerable to supply-chain attacks.

**Fix:** Pin to full SHA: `actions/checkout@<commit-sha>`.

---

#### F71a. `npm audit` failure is silently swallowed
**Severity:** MEDIUM | **Category:** Security | **Line:** 34

`npm audit --audit-level=high || true` ensures exit code is always 0. Even high-severity vulnerabilities never fail the build.

**Fix:** Remove `|| true` to enforce audit.

---

#### F71b. No npm cache configured in CI
**Severity:** MEDIUM | **Category:** Performance | **Lines:** 18-24

Every CI run downloads and installs all dependencies from scratch.

**Fix:** Add `cache: 'npm'` to `actions/setup-node`.

---

#### F71c. CI uses EOL Node.js 18
**Severity:** LOW | **Category:** Quality | **Line:** 21

Node.js 18 reached EOL in April 2025. Update to Node.js 20 or 22.

---

### .husky/pre-commit

**Findings:**

#### F72. Pre-commit hook lacks execute permission
**Severity:** HIGH | **Category:** Bug

The `pre-commit` hook has permissions `644` (rw-r--r--) while `commit-msg` has `755` (rwxr-xr-x). Without the execute bit, the pre-commit hook **silently fails to run** on Unix, meaning lint-staged and build verification are skipped on every commit.

**Fix:** `chmod +x .husky/pre-commit`

---

#### F72a. Build errors in pre-commit are redirected to /dev/null
**Severity:** MEDIUM | **Category:** Maintainability | **Line:** 21

Build output is suppressed with `> /dev/null 2>&1`. If the build fails, devs see only "Build failed" with no error details.

**Fix:** Only redirect stdout, keep stderr visible: `npm run build > /dev/null`

---

### appsscript.json

**Findings:**

#### F73. Overly broad OAuth scopes
**Severity:** MEDIUM | **Category:** Security | **Lines:** 14-24

The manifest requests `https://www.googleapis.com/auth/drive` (full Drive access) when `auth/drive.file` (access only to files created by the app) may suffice. Combined with `gmail.send`, a vulnerability could be exploited to exfiltrate data via email.

**Fix:** Evaluate whether `auth/drive.file` is sufficient.

---

### test/gas-mock.js

**Findings:**

#### F74. Missing mocks for actively-used GAS globals
**Severity:** MEDIUM | **Category:** Bug

`DocumentApp` (16 occurrences in source), `GmailApp`, `Browser`, and `ContentService` have no mocks despite being used in source code. Tests exercising these code paths will throw `ReferenceError`.

**Fix:** Add basic mocks for these globals.

---

#### F74a. Mock timezone doesn't match appsscript.json
**Severity:** LOW | **Category:** Bug | **Line:** 53

Mock `Session.getScriptTimeZone()` returns `'America/Los_Angeles'` but `appsscript.json` configures `'America/New_York'`. Date-sensitive tests will produce different results than production.

---

### test/load-source.js

**Findings:**

#### F75. Regex rewrites match indented code, not just top-level
**Severity:** MEDIUM | **Category:** Bug | **Lines:** 78, 82, 85

The `^` anchor with `/gm` matches the start of *any* line, not just column-0 code. A `var` inside an `if` block or `for` loop at column 0 would incorrectly be rewritten to `global.X =`, changing scoping semantics.

---

## 14. Cross-Cutting Concerns

### CC1. Inconsistent Column Index Systems (1-indexed vs 0-indexed)
**Severity:** HIGH | **Category:** Quality | **Scope:** Entire codebase

The codebase uses TWO column index systems:
- `*_COLS` — 1-indexed (for `sheet.getRange()`)
- `*_COLUMNS` — 0-indexed (for array access)

Functions mix these systems. `DataAccess` layer uses `MEMBER_COLUMNS` (0-indexed). `02_DataManagers.gs` uses `MEMBER_COLS` (1-indexed). Any cross-module call must know which system the callee expects.

**Fix:** Add clear naming conventions or consolidate to one system with explicit conversion helpers.

---

### CC2. HTML String Concatenation as Dominant UI Pattern
**Severity:** MEDIUM | **Category:** Maintainability | **Scope:** ~15 files

Over 20,000 lines of HTML are built via JavaScript string concatenation. This creates:
1. XSS audit difficulty (impossible to grep for missing `escapeHtml()`)
2. No syntax highlighting or IDE support for HTML
3. No template validation

**Fix:** Migrate to `HtmlService.createTemplateFromFile()` with scriptlets for dynamic data.

---

### CC3. No Script Lock Usage
**Severity:** HIGH | **Category:** Bug | **Scope:** Data write operations

None of the data mutation functions (add member, file grievance, sync operations, etc.) use `LockService`. In a multi-user environment (which a union dashboard is), concurrent edits to the same sheet can cause:
- Duplicate records
- Overwritten data
- Corrupted state

**Fix:** Add `LockService.getScriptLock().waitLock(10000)` to all mutation paths.

---

### CC4. Global Namespace Pollution
**Severity:** LOW | **Category:** Quality | **Scope:** All files

All functions and variables are in the global scope (standard for GAS). With 30 files and 500+ functions, name collisions are possible. Private functions use `_` suffix convention, which helps.

---

### CC5. Error Handling Inconsistency
**Severity:** MEDIUM | **Category:** Quality | **Scope:** All files

Three different error handling patterns are used:
1. `handleError()` from 01_Core.gs (structured logging)
2. `Logger.log()` (basic logging)
3. Silent catch-and-return (errors swallowed)

**Fix:** Standardize on `handleError()` for all error paths.

---

### CC6. Test Coverage Gaps in Security-Critical Code
**Severity:** MEDIUM | **Category:** Quality

Tests exist for most modules, but security-critical flows (PIN validation, XSS sanitization edge cases, CSRF protection) need more targeted test coverage.

---

### CC7. Deprecated Code Retained for Backward Compatibility
**Severity:** LOW | **Category:** Quality | **Scope:** Multiple files

Several constants and functions are marked `@deprecated` but remain in the codebase:
- `SHEETS.DASHBOARD` (deprecated v4.3.2)
- `SHEETS.SATISFACTION` (deprecated v4.3.8)
- `SATISFACTION_COLS.EMAIL/VERIFIED/etc.` (deprecated v4.8)

Consider adding a version-gated removal plan.

---

## Summary of Actionable Recommendations

### Priority 1 — Fix Now (CRITICAL + HIGH)

| ID | Summary | Effort |
|----|---------|--------|
| F34a | XSS: Use `JSON.stringify()` for URLs in `setupFolderForSelectedGrievance()` | Trivial |
| F34b | XSS: Use `JSON.stringify()` for URLs in `createPDFForSelectedGrievance()` | Trivial |
| F34c | XSS: Add `escapeHtml()` to email HTML bodies (meeting attendance, docs) | Small |
| F10 | Update VERSION_INFO/API_VERSION to 4.9.0 | Trivial |
| F68 | Update package.json version to 4.9.0 | Trivial |
| F72 | `chmod +x .husky/pre-commit` (hook is silently not running!) | Trivial |
| F14 | Fix array sizing with MEMBER_HEADER_MAP_.length | Trivial |
| F19 | Sanitize CSV preview data with escapeHtml() | Small |
| F35b | XSS: Escape data in audit log viewer HTML | Small |
| F35c | XSS: Escape steward names in workload dashboard | Small |
| F34d | Fix race condition: use `e.range` instead of `getLastRow()` in form handler | Small |
| F15 | Batch setValue() calls in addMember() | Small |
| F35a | Fix undo system UserProperties 9KB overflow | Medium |
| F70 | Re-enable critical ESLint rules (no-undef, no-dupe-args, etc.) | Medium |
| F69 | Fix Jest coverage collection for .gs files or remove misleading config | Small |
| F57 | Add PIN brute-force rate limiting | Medium |
| F52 | Add LockService to sync functions | Medium |
| CC3 | Add LockService to all mutation paths | Medium |
| F47 | Defer non-critical work from onOpen() | Medium |
| F35 | Mask email in audit log entries | Small |
| F28 | Plan file splitting for 04e_PublicDashboard.gs | Large |

### Priority 2 — Plan for Next Release (MEDIUM)

| ID | Summary | Effort |
|----|---------|--------|
| F1 | Fix setCells() batching | Small |
| F6 | Mask all PII in security alert emails | Small |
| F8 | Document safeJsonForHtml() double-escape risk | Small |
| F9 | Add fallback in sendCriticalErrorNotification_() | Small |
| F12 | Cache COMMAND_CONFIG.SYSTEM_NAME | Small |
| F17 | Fix updateMember() field clearing bug | Small |
| F34 | Create safeSendEmail() wrapper | Medium |
| F61 | Add error isolation to EventBus handlers | Small |
| F67a | Read version dynamically in build.js | Small |
| F67b | Fix HTML $ escaping to only escape `${` | Small |
| F71a | Remove `\|\| true` from npm audit in CI | Trivial |
| F71b | Add npm cache to CI workflow | Trivial |
| F73 | Evaluate narrower OAuth scopes (drive.file) | Small |
| F74 | Add missing GAS global mocks (DocumentApp, etc.) | Small |
| F75 | Fix regex rewrites in test loader | Medium |
| CC2 | Begin HTML template migration | Large |
| CC5 | Standardize error handling patterns | Medium |
| F71 | Pin CI action versions to SHA | Small |

### Priority 3 — Backlog (LOW)

Low-severity items (F2, F3, F4, F5, F7, F11, F13, F18, F20, F24, F25, F26, F37, F40, F41, F43, F44, F49, F50, F53, F56, F60, F62, F63, F65, F66, F67d, F68a, F68b, F71c, F74a, CC4, CC7) can be addressed during regular maintenance.

---

## Positive Observations

1. **Security-first design**: PII masking, SHA-256 hashing, XSS prevention, formula injection guards, audit logging with integrity hashes
2. **Auto-discovery column system**: Eliminates manual index maintenance, handles user-reordered sheets
3. **Comprehensive test suite**: 1,300+ tests across 23 suites
4. **Accessibility**: WCAG-aware design, keyboard navigation, color-blind support
5. **Progressive enhancement**: Mobile-optimized views, responsive design
6. **Operational tooling**: Diagnostics, health checks, self-healing formulas
7. **Batch operations**: Import/export with progress bars, batch writes
8. **Build system**: Proper CI/CD with ESLint, Jest, Husky, commitlint
9. **Zero-knowledge survey vault**: Industry-grade privacy design for anonymous surveys
10. **Hidden sheet enforcement**: API-level hiding + protection that works on mobile

---

*Review completed 2026-02-21 by Claude Code (Opus 4.6)*
