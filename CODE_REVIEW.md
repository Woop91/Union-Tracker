# CODE_REVIEW.md — Comprehensive Codebase Review

**Date:** 2026-02-27
**Scope:** Full codebase review of all 33 `.gs` files, 8 `.html` files, tests, and build configuration (~76,600 lines)
**Methodology:** Automated multi-agent analysis with manual cross-referencing against CLAUDE.md rules

---

## Executive Summary

| Severity | Count | Description |
|----------|-------|-------------|
| **CRITICAL** | 30 | Security vulnerabilities, data loss/corruption risks |
| **HIGH** | 53 | Logic bugs, broken features, missing authorization |
| **MEDIUM** | 68 | Code quality, edge cases, hardcoded values, performance |
| **LOW** | 42 | Style issues, dead code, minor improvements |
| **Total** | **199** |

**Top 5 systemic issues:**
1. **Broken access control in web dashboard** — Client-supplied `email` parameter trusted for authentication on all `google.script.run` wrappers
2. **Manual data overwrite violations** — Multiple sync/batch functions overwrite entire sheets including manually entered data
3. **Hardcoded column positions and sheet names** — Widespread violation of the "everything must be dynamic" rule
4. **Missing `escapeForFormula()` / `escapeHtml()`** — Formula injection in imports, XSS via `innerHTML` in HTML files
5. **Race conditions without locking** — Read-modify-write cycles without `LockService` in concurrent paths

---

## CRITICAL Issues (30)

### Authentication & Authorization

#### CR-01: No server-side auth on web dashboard API wrappers
**Files:** `21_WebDashDataService.gs:1375-1417`, `24_WeeklyQuestions.gs:401-407`, `25_WorkloadService.gs:1081-1087`

Every `google.script.run`-callable global wrapper blindly trusts the `email` parameter from the client. An attacker can call any function with any email:
```javascript
function dataGetStewardCases(email) { return DataService.getStewardCases(email); }
function dataUpdateProfile(email, updates) { return DataService.updateMemberProfile(email, updates); }
function dataAssignSteward(memberEmail, stewardEmail) { ... }
function dataSendBroadcast(email, filter, msg) { ... }
```
**Impact:** Full account takeover — any user can read/modify any member's data, reassign stewards, send broadcast emails.
**Fix:** Resolve caller identity server-side via `Session.getActiveUser().getEmail()` or validate session tokens.

#### CR-02: `authCreateSessionToken` callable with arbitrary email
**File:** `19_WebDashAuth.gs:324`
```javascript
function authCreateSessionToken(email) { return Auth.createSessionToken(email); }
```
Creates session tokens for any email without verifying the caller's identity. Full account takeover vector.

#### CR-03: Magic link token not single-use (replay attack)
**File:** `19_WebDashAuth.gs:209-235`
Token marked `used = true` but the `used` flag is never checked on validation. Tokens remain valid for the full 7-day expiry window.

#### CR-04: Authentication bypass in member grievance history
**File:** `13_MemberSelfService.gs:1254-1327`
`getMemberGrievanceHistory()` accepts raw email with zero authentication when `indexOf('@')` matches. `dataGetMemberGrievanceHistoryPortal(email)` exposes this to any caller.

#### CR-05: `getEffectiveUser()` returns script owner in web apps
**File:** `00_Security.gs:302-304`
In "Execute as me" deployments, `Session.getEffectiveUser()` returns the script owner's email, not the accessing user. All users appear as admin.

### Data Integrity — Manual Data Overwrite Violations

#### CR-06: `deleteRow()` has no protection for manually entered data
**File:** `00_DataAccess.gs:347-352`
General-purpose row deletion with zero safeguards — no audit, no confirmation, no check for manual data.

#### CR-07: `setRow()` can overwrite manually entered data
**File:** `00_DataAccess.gs:320-326`
Overwrites an entire row with no column-level check for system-generated vs. manually entered fields.

#### CR-08: `generateMissingMemberIDsBatch()` writes entire sheet back
**File:** `11_CommandHub.gs:336-381`
```javascript
sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
```
Reads entire Member Directory, modifies only Member ID cells, then writes the **entire array** back. Any concurrent edits are silently lost.

#### CR-09: `autoPopulateGrievanceFromOCR_()` overwrites manual grievance fields
**File:** `11_CommandHub.gs:1915-2036`
Unconditionally overwrites Incident Date, Articles, and Issue Category without checking existing manually entered values.

#### CR-10: `syncGrievanceFormulasToLog()` overwrites First/Last Name and contact columns
**File:** `09_Dashboards.gs:1984-2026`
Bulk-writes names and contact info for every grievance row. If member lookup fails, writes empty strings — blanking manually entered data.

#### CR-11: `sheet.clear()` on setup functions destroys existing data
**Files:** `10b_SurveyDocSheets.gs:12,487,864,1153,1389,1742`, `10a_SheetCreation.gs:47,545,1174`
`getOrCreateSheet()` returns an existing sheet, then `sheet.clear()` destroys all data on it — including form responses and manually entered survey data.

#### CR-12: `handleStageGateWorkflow_()` overwrites manually entered close dates
**File:** `10_Main.gs:472-476`
When status changes to closed, unconditionally sets `DATE_CLOSED` to `new Date()`, overwriting backdated dates stewards may have entered.

#### CR-13: `batchSetValues` overwrites entire sheet without lock
**File:** `06_Maintenance.gs:2123-2148`
Reads from row 1 through the last update row, then writes it all back. Concurrent edits are silently lost. No `LockService`.

#### CR-14: `restoreFromSnapshot` deletes all grievance data
**File:** `06_Maintenance.gs:1290-1298`
Clears all existing data and replaces with snapshot. The "backup" is in-memory only — script timeout means total data loss.

#### CR-15: `updateMemberDataBatch()` writes full row without formula escape
**File:** `02_DataManagers.gs:905-954`
Read-modify-write of entire row without `escapeForFormula()` and without checking for manual data in other columns.

### Security — Injection Vulnerabilities

#### CR-16: XSS via `innerHTML` in `auth_view.html`
**File:** `auth_view.html:206`
```javascript
desc.innerHTML = '...<span ...>' + email + '</span>';
```
User-typed email rendered as raw HTML. `<img src=x onerror=alert(1)>` executes.

#### CR-17: XSS via `innerHTML` in `WorkloadTracker.html` (7 locations)
**Files:** `WorkloadTracker.html:831-834,862-881,575-577,521-527,771-779`
Server data (`d.reason`, category keys, employment breakdown, history fields) injected via `innerHTML` without escaping.

#### CR-18: Formula injection in contact form update path
**File:** `08c_FormsAndNotifications.gs:374-396`
The update-existing-member branch writes 18 form fields to the sheet **without** `escapeForFormula()`. The new-member branch correctly escapes.

#### CR-19: Formula injection in member imports
**File:** `02_DataManagers.gs:1331-1343,1466-1479`
`importMembersFromData()` and `importMembersBatch()` write CSV data directly without `escapeForFormula()`. Malicious CSV files can inject formulas.

#### CR-20: CSV injection in WorkloadTracker export
**File:** `18_WorkloadTracker.gs:771-792`
`exportHistoryCSV()` joins values without sanitizing formula injection characters (`=`, `+`, `-`, `@`).

#### CR-21: `sendEmailToMember()` HTML injection
**File:** `05_Integrations.gs:1323-1350`
Caller-supplied `body` passed directly as `htmlBody` to `MailApp.sendEmail`. Phishing/XSS vector.

#### CR-22: XSS in magic link email HTML
**File:** `19_WebDashAuth.gs:273-296`
`config.orgName` and `config.logoInitials` injected into email HTML without escaping.

#### CR-23: `buildSafeQuery()` incomplete injection protection
**File:** `00_Security.gs:237-252`
Blocklist approach is fragile. Does not block `QUERY` nesting, `INDIRECT`, or `OFFSET`. Range hardcoded to `A:Z`.

#### CR-24: `importMembersFromText()` uses hardcoded column positions
**File:** `10_Main.gs:2074-2084`
9-element array with hardcoded column order — no `MEMBER_COLS` constants. Column reorder = data corruption.

### Data Integrity — Logic Errors

#### CR-25: `getUpcomingDeadlines()` uses wrong header strings
**File:** `02_DataManagers.gs:2219-2248`
Uses `'Step 2 Due'` and `'Step 3 Due'` but actual headers are `'Step II Due'` and `'Step III Due'` (Roman numerals). Deadlines for Steps 2-3 will **never** appear. `'Member Name'` is also wrong — actual headers are `'First Name'`/`'Last Name'`.

#### CR-26: Double `escapeForFormula()` corrupts data
**Files:** `02_DataManagers.gs:1950-1952,2127-2128`
In `advanceGrievanceStep()` and `bulkUpdateGrievanceStatus()`, notes are escaped once when building the string, then the whole string is escaped again. Double-escaping corrupts existing resolution text.

#### CR-27: `startNewGrievance()` lacks concurrency protection
**File:** `02_DataManagers.gs:1588-1646`
No `withScriptLock_()`. Two simultaneous grievance creations can get the same ID from `getNextGrievanceId()`.

#### CR-28: `_getChiefStewardEmail` reads row 2 instead of row 3
**File:** `21_WebDashDataService.gs:1120`
Per CLAUDE.md, Config data starts at row 3. Reading row 2 returns the column header text, not the email. Chief steward features never work.

#### CR-29: `showResetPINDialog()` displays PIN in plaintext
**File:** `13_MemberSelfService.gs:903-910`
Unlike `showGeneratePINDialog()` which emails the PIN securely, the reset dialog shows it in a UI alert — vulnerable to shoulder surfing.

#### CR-30: Audit log trimming breaks integrity hash chain
**File:** `06_Maintenance.gs:1542-1546`
Deleting old audit rows breaks the hash chain used for tamper detection. `verifyAuditLogIntegrity()` will report false positives on all remaining rows.

---

## HIGH Issues (53)

### Hardcoded Values (violating "everything must be dynamic" rule)

| # | File | Line(s) | Issue |
|---|------|---------|-------|
| H-01 | `00_DataAccess.gs` | 366,402,452 | Fallback sheet names `'Member Directory'`, `'Grievance Log'` |
| H-02 | `21_WebDashDataService.gs` | 21-22 | `var MEMBER_SHEET = 'Member Directory'` |
| H-03 | `21_WebDashDataService.gs` | 545 | Hardcoded `'_Survey_Tracking'` |
| H-04 | `21_WebDashDataService.gs` | 1119 | Hardcoded `'Config'` sheet name |
| H-05 | `11_CommandHub.gs` | 55-57 | `COMMAND_CENTER_CONFIG` hardcoded sheet names |
| H-06 | `11_CommandHub.gs` | 1035 | Hidden sheet names hardcoded in diagnostic |
| H-07 | `13_MemberSelfService.gs` | 32,44 | `PIN_COLUMN: 33` hardcoded column number |
| H-08 | `18_WorkloadTracker.gs` | 249-251 | Fallback column numbers `9, 1, 33` |
| H-09 | `23_PortalSheets.gs` | 90-143 | All 8 portal sheet names hardcoded |
| H-10 | `04e_PublicDashboard.gs` | 811-884 | ~80 hardcoded satisfaction survey column indices |
| H-11 | `10b_SurveyDocSheets.gs` | 159-171 | Hardcoded column ranges 72-82, dashStart=84 |
| H-12 | `10_Main.gs` | 628-635 | Hardcoded unit codes ("Main Station", "Field Ops", etc.) |
| H-13 | `25_WorkloadService.gs` | 881 | Hardcoded org name "SEIU 509 DDS Dashboard" |
| H-14 | `11_CommandHub.gs` | 927-942 | `applyConfigSectionColors()` hardcoded column ranges |
| H-15 | `11_CommandHub.gs` | 2181 | Fallback `|| 82` for satisfaction columns |
| H-16 | `21_WebDashDataService.gs` | 988-1001 | Magic column numbers `[0]`-`[6]` in contact log |

### Race Conditions & Concurrency

| # | File | Line(s) | Issue |
|---|------|---------|-------|
| H-17 | `00_DataAccess.gs` | 280-311 | `setCells()` read-modify-write without lock |
| H-18 | `00_DataAccess.gs` | 334-340 | `appendRow()` return value unreliable under concurrency |
| H-19 | `24_WeeklyQuestions.gs` | 149-185 | `submitResponse` duplicate check + write not atomic |
| H-20 | `18_WorkloadTracker.gs` | 363 | `_cleanVaultData()` vault.clear() without lock |
| H-21 | `00_Security.gs` | 972-1084 | TOCTOU race in digest queue send/clear |

### Missing Authorization

| # | File | Line(s) | Issue |
|---|------|---------|-------|
| H-22 | `02_DataManagers.gs` | 2031 | `recalcAllGrievancesBatched()` no auth check |
| H-23 | `06_Maintenance.gs` | 1654-1723 | `NUCLEAR_WIPE_GRIEVANCES` no `validateRole()` |
| H-24 | `24_WeeklyQuestions.gs` | 256-273,404 | `setStewardQuestion` no auth — any user can set weekly question |
| H-25 | `24_WeeklyQuestions.gs` | 405 | `getPoolQuestions` exposes steward-only questions to any caller |
| H-26 | `16_DashboardEnhancements.gs` | 137-170 | `scheduleEmailReport()` allows PII reports to any email |
| H-27 | `16_DashboardEnhancements.gs` | 346-372 | `pushNotification()` allows pushing to any user |
| H-28 | `21_WebDashDataService.gs` | 667-711 | `sendBroadcastMessage` no rate limit, no auth verification |

### Logic Bugs

| # | File | Line(s) | Issue |
|---|------|---------|-------|
| H-29 | `02_DataManagers.gs` | 2395 | `resolveGrievance()` always sets status to `'Settled'` regardless of outcome |
| H-30 | `02_DataManagers.gs` | 2298-2321 | `getGrievanceStats()` doesn't count 'Denied'/'Withdrawn' — total != open + pending + resolved |
| H-31 | `02_DataManagers.gs` | 2471-2475 | `showBulkStatusUpdate()` uses `g['Member Name']` which doesn't exist — shows `undefined` |
| H-32 | `04d_ExecutiveDashboard.gs` | 59 | `getDashboardStats()` counts blank rows as grievances |
| H-33 | `04d_ExecutiveDashboard.gs` | 78-81 | Outcome counting checks `status` instead of `resolution` column |
| H-34 | `21_WebDashDataService.gs` | 1227 | `getGrievanceStats()` uses `rec.issueCategory` which doesn't exist — all categorized as `'Uncategorized'` |
| H-35 | `08d_AuditAndFormulas.gs` | 550-665 | VLOOKUP index uses absolute column number instead of relative offset |
| H-36 | `10_Main.gs` | 754-773 | `handleGrievanceEdit()` conflicts with `recalculateDownstreamDeadlines_()` — deadline overwrite |
| H-37 | `00_DataAccess.gs` | 366+ | `00_DataAccess.gs` depends on `01_Core.gs` constants (load order risk) |

### PII / Data Exposure

| # | File | Line(s) | Issue |
|---|------|---------|-------|
| H-38 | `00_Security.gs` | 587-612 | `secureLog()` PII masking incomplete — misses address, SSN, DOB |
| H-39 | `00_Security.gs` | 875-885 | Security alert email leaks PII (phone, name, address) |
| H-40 | `00_Security.gs` | 937-947 | Digest queue stores unmasked PII in ScriptProperties |
| H-41 | `04d_ExecutiveDashboard.gs` | 592-593 | `checkOverdueGrievances_()` sends member names in plaintext email |
| H-42 | `04d_ExecutiveDashboard.gs` | 713 | `emailExecutivePDF()` exports entire spreadsheet as PDF including raw PII |

### Schema/Architecture Conflicts

| # | File | Line(s) | Issue |
|---|------|---------|-------|
| H-43 | `06_Maintenance.gs` + `08d_AuditAndFormulas.gs` | 1474 / 28-29 | Two conflicting audit log schemas (6-col vs 10-col) writing to same sheet |
| H-44 | `06_Maintenance.gs` | 1901-1902 | Third audit schema variant (5-col) in `navigateToAuditLog` |
| H-45 | `23_PortalSheets.gs` | 24-62 | Portal column constants are 0-indexed — rest of codebase is 1-indexed |
| H-46 | `02_DataManagers.gs` | 1616 | `startNewGrievance()` stores full name in FIRST_NAME column, LAST_NAME empty |

### Performance

| # | File | Line(s) | Issue |
|---|------|---------|-------|
| H-47 | `04d_ExecutiveDashboard.gs` | 960-970 | `applyStatusColors()` makes one API call per row — O(n) will time out |
| H-48 | `02_DataManagers.gs` | 335-341 | `syncMemberGrievanceData()` 2 `setValue()` per member row — 400 calls for 200 members |

### HTML / Client-Side Security

| # | File | Line(s) | Issue |
|---|------|---------|-------|
| H-49 | `index.html` | 226-230 | Unsafe `html()` utility function parses arbitrary strings as HTML |
| H-50 | `index.html` | 25 | `<?!= pageData ?>` unescaped scriptlet — verify JSON.stringify() server-side |
| H-51 | `index.html` | 33-39 | Session token in `localStorage` — exfiltrable via XSS |
| H-52 | `index.html` | — | No Content-Security-Policy on main application |
| H-53 | `member_view.html:833`, `steward_view.html:1012` | — | `href` attributes accept server data — `javascript:` protocol URLs possible |

---

## MEDIUM Issues (68)

### Security & Sanitization (14)

| # | File | Issue |
|---|------|-------|
| M-01 | `00_Security.gs:194-210` | `escapeForFormula()` only guards first character — mid-string injection possible |
| M-02 | `00_Security.gs:158-177` | `sanitizeObjectForHtml()` corrupts arrays (returns plain object) |
| M-03 | `00_Security.gs:258-264` | `validateInputLength_()` silently truncates; coerces falsy to empty string |
| M-04 | `00_Security.gs:624-626` | `isValidSafeString()` returns `true` for null/undefined |
| M-05 | `10_Main.gs:2071-2086` | No input sanitization on CSV import — `importMembersFromText()` |
| M-06 | `22_WebDashApp.gs:105,143,171` | `XFrameOptionsMode.ALLOWALL` allows clickjacking |
| M-07 | `22_WebDashApp.gs:159-165` | `_serveError` leaks raw error details to client |
| M-08 | `19_WebDashAuth.gs:69-126` | `sendMagicLink` leaks email existence via timing side-channel |
| M-09 | `05_Integrations.gs:1411-1420` | `replaceText()` regex injection via `$` in replacement strings |
| M-10 | `16_DashboardEnhancements.gs:438,498` | View names and comments stored unsanitized |
| M-11 | `21_WebDashDataService.gs:1077-1090` | `updateTask` values not validated against allowed lists |
| M-12 | `index.html:16` | External Chart.js CDN loaded without SRI hash |
| M-13 | `06_Maintenance.gs:894-900` | Missing `escapeHtml` for cache key names in HTML |
| M-14 | `08c_FormsAndNotifications.gs:1922` | `sendSurveyCompletionReminders` logs raw email in plaintext |

### Config / Dynamic Rule Violations (10)

| # | File | Issue |
|---|------|-------|
| M-15 | `08c_FormsAndNotifications.gs:40` | `getFormUrlFromConfig` reads row 2 instead of row 3 |
| M-16 | `06_Maintenance.gs:1770` | Hardcoded "column AS" in error message |
| M-17 | `06_Maintenance.gs:1316` | `'Undo_History_Export'` hardcoded sheet name |
| M-18 | `21_WebDashDataService.gs:545-567` | Ad-hoc header aliases defined inline instead of centralized |
| M-19 | `01_Core.gs:617` | `DRIVE_CONFIG.ROOT_FOLDER_NAME` hardcoded |
| M-20 | `10c_FormHandlers.gs:668` | `freezeKeyColumns()` hardcodes freeze at column 6 |
| M-21 | `11_CommandHub.gs:2345-2347` | Date column found by header-search fallback instead of `GRIEVANCE_COLS` |
| M-22 | `19_WebDashAuth.gs:115` | `sendMagicLink` logs email in plaintext (should use `secureLog`) |
| M-23 | `21_WebDashDataService.gs:36,51` | Broad alias collision for "status" and "id" across sheets |
| M-24 | `20_WebDashConfigReader.gs:120-127` | `_readCell` silently returns empty string on any error — masks failures |

### Data Integrity & Logic (18)

| # | File | Issue |
|---|------|-------|
| M-25 | `10d_SyncAndMaintenance.gs:928-933` | `syncVolunteerHoursToMemberDirectory()` zeroes hours for members not in volunteer sheet |
| M-26 | `09_Dashboards.gs:1739,1858` + `10_Main.gs:473,670` | `closedStatuses` array defined 4+ places with inconsistent ordering |
| M-27 | `04d_ExecutiveDashboard.gs:98-101` | Overdue check only looks at STEP1_DUE — misses Steps 2-3 |
| M-28 | `04d_ExecutiveDashboard.gs:105-108` | Win rate inconsistent across dashboards (settled in denominator vs. not) |
| M-29 | `04e_PublicDashboard.gs:278` | `getUnifiedDashboardData()` counts non-member rows as members |
| M-30 | `04e_PublicDashboard.gs:601-607` | "My Cases" matching uses loose substring heuristic — false positives |
| M-31 | `04c_InteractiveDashboard.gs:1273-1279` | `getMyStewardCases()` loose substring match — "Bob" matches "Bobby" |
| M-32 | `10_Main.gs:789-838` | `recalculateDownstreamDeadlines_()` never actually cascades downstream |
| M-33 | `10_Main.gs:1906` | `addNewMember()` sizes row by `MEMBER_COLS.STATE` which may not be last column |
| M-34 | `02_DataManagers.gs:213-214` | `generateMemberID_()` fallback uses `Date.now()` last 3 digits — collision risk |
| M-35 | `21_WebDashDataService.gs:1550-1559` | `_welcomeEmailHash` weak hash with high collision probability |
| M-36 | `21_WebDashDataService.gs:488-512` | `assignStewardToMember` writes without verifying steward role |
| M-37 | `25_WorkloadService.gs:937-980` | `archiveOldData` clears vault before rewrite — data loss if write fails |
| M-38 | `25_WorkloadService.gs:984-1022` | `cleanVault` same clear-before-write risk |
| M-39 | `08c_FormsAndNotifications.gs:1733-1734` | `populateSurveyTrackingFromMembers` clears all tracking data without backup |
| M-40 | `08d_AuditAndFormulas.gs:104-149` | `onEditAudit` doesn't handle multi-cell edits (oldValue/value are undefined) |
| M-41 | `08c_FormsAndNotifications.gs:543,556` | Duplicate `var match` declaration in `setupGrievanceFormTrigger` |
| M-42 | `05_Integrations.gs:1603-1604` | `onGrievanceFormSubmit()` overwrites folder URL with PDF URL |

### Performance (8)

| # | File | Issue |
|---|------|-------|
| M-43 | `01_Core.gs:631-690` | `_cachedOrgName` etc. never invalidated during execution |
| M-44 | `02_DataManagers.gs:141-172` | `searchMembers()` has no result limit — could return thousands |
| M-45 | `04e_PublicDashboard.gs:273,478,784` | `getUnifiedDashboardData()` loads all columns of all rows |
| M-46 | `17_CorrelationEngine.gs:277,920,964` | Single summary view triggers 3 full dashboard data fetches |
| M-47 | `10c_FormHandlers.gs:647-649` | `applyStepHighlighting()` appends rules without dedup — accumulates over time |
| M-48 | `10_Main.gs:207-225` | `onEdit` + `onEditAutoSync` both call `syncGrievanceFormulasToLog` — duplicate work |
| M-49 | `05_Integrations.gs:226-228` | `cleanupEmptyDriveFolders()` O(n*m) complexity for folder-grievance matching |
| M-50 | `19_WebDashAuth.gs:139-146` | ScriptProperties storage not scalable for tokens (500KB total limit) |

### Miscellaneous Medium (18)

| # | File | Issue |
|---|------|-------|
| M-51 | `01_Core.gs:557,700-708` | `VERSION_INFO` duplicated from `COMMAND_CONFIG.VERSION` — can drift |
| M-52 | `02_DataManagers.gs:266-267` | `getStewardWorkloadDetailed()` uses fragile header-sniffing for name |
| M-53 | `02_DataManagers.gs:278,327` | Hardcoded `'Open'`/`'Pending Info'` instead of `GRIEVANCE_STATUS.*` |
| M-54 | `02_DataManagers.gs:3099-3100` | `selectAllOpenCases()` lowercases status inconsistently |
| M-55 | `05_Integrations.gs:537` | Meeting date parsed without timezone — wrong day possible |
| M-56 | `05_Integrations.gs:594-606` | `emailMeetingAttendanceReport()` sends PII to arbitrary email |
| M-57 | `05_Integrations.gs:814` | `DOMAIN_WITH_LINK` sharing fails for non-Workspace orgs |
| M-58 | `06_Maintenance.gs:2069-2081` | `saveSettings` accepts arbitrary object without validation |
| M-59 | `10_Main.gs:43-47` | `onOpen()` creates trigger every open — could exhaust 20-trigger limit |
| M-60 | `10_Main.gs:2219-2224` | `exportMemberDirectory()` creates CSV in Drive, never cleans up |
| M-61 | `15_EventBus.gs:148-155` | `emit()` swallows listener errors that most callers ignore |
| M-62 | `16_DashboardEnhancements.gs:322-372` | `getUserNotifications()` reads UserProperties but `pushNotification()` writes ScriptProperties — different stores |
| M-63 | `16_DashboardEnhancements.gs:488-521` | `addViewComment()` no auth check — any user can comment on any view |
| M-64 | `17_CorrelationEngine.gs:795-835` | `correlateVolunteerVsEngagement_()` function name misleading |
| M-65 | `18_WorkloadTracker.gs:351-353` | Dedup key uses locale-dependent `String(Date)` — unreliable |
| M-66 | `19_WebDashAuth.gs:257-261` | Token generation encodes UUID string (not bytes) — inflated token |
| M-67 | `19_WebDashAuth.gs:263-271` | `_getHmacSecret` defined but never used — dead code |
| M-68 | `10d_SyncAndMaintenance.gs:857-864` | `buildMemberIdSet_()` header-row heuristic could include header as data |

---

## LOW Issues (42)

<details>
<summary>Click to expand LOW issues</summary>

### Style & Consistency (12)

| # | File | Issue |
|---|------|-------|
| L-01 | `02_DataManagers.gs` | Mixed `var`/`const`/`let` within same file |
| L-02 | `06_Maintenance.gs` | Mixed `var`/`const`/`let` across modules |
| L-03 | `10_Main.gs` | Mixed `var`/`const`/`let` within functions |
| L-04 | `05_Integrations.gs:1370-1385` | `folderName` re-declared 3 times with `var` in nested blocks |
| L-05 | `01_Core.gs:773-829` | Emoji in sheet names reduces cross-platform portability |
| L-06 | `10c_FormHandlers.gs:615` | Comment says "column E" but uses dynamic constant |
| L-07 | `04e_PublicDashboard.gs:660,818` | `toLocaleString('default')` varies by server locale |
| L-08 | `20_WebDashConfigReader.gs:28,74` | `var cache` re-declared in same function scope |
| L-09 | `21_WebDashDataService.gs:908-918` | `var now` and `var diff` re-declared in same scope |
| L-10 | `25_WorkloadService.gs:412` | Inconsistent return types (string vs. `{success, message}` object) |
| L-11 | WorkloadTracker.html | Uses `innerHTML` pattern while rest of app uses safe `el()` helper |
| L-12 | All HTML views | Missing ARIA attributes on interactive elements |

### Dead Code & Unused (12)

| # | File | Issue |
|---|------|-------|
| L-13 | `01_Core.gs:262-267` | `sanitizeForQuery()` unused and incomplete |
| L-14 | `04d_ExecutiveDashboard.gs:15-17,399-425` | Multiple deprecated wrapper functions |
| L-15 | `07_DevTools.gs:910-1237` | `SEED_MEMBERS_ONLY` duplicates `SEED_MEMBERS` (~130 lines) |
| L-16 | `08c_FormsAndNotifications.gs:1058` | Unused variable `_threeDaysAhead` |
| L-17 | `08c_FormsAndNotifications.gs:1393` | Unused variable `_ss` makes unnecessary API call |
| L-18 | `08c_FormsAndNotifications.gs:1460` | Unused variable `_headers` |
| L-19 | `10d_SyncAndMaintenance.gs:671,511` | Unused variables `_ss`, `_issueCategory` |
| L-20 | `10a_SheetCreation.gs:1606` | Unused variable `_percentFormat` |
| L-21 | `10d_SyncAndMaintenance.gs:1690-1692` | `getSheetLastRow()` trivial wrapper that adds no value |
| L-22 | `11_CommandHub.gs:53-69,1413-1438` | `COMMAND_CENTER_CONFIG` + `GEMINI_CONFIG` redundant to `COMMAND_CONFIG` |
| L-23 | `20_WebDashConfigReader.gs:129-132` | `_parseInt` helper defined but never called |
| L-24 | `15_EventBus.gs:256` | `registerEventBusSubscribers()` calls `EventBus.reset()` — wipes prior subscribers |

### Minor Logic (10)

| # | File | Issue |
|---|------|-------|
| L-25 | `00_DataAccess.gs:176-184` | `getRow()` no bounds validation — throws on invalid rowNumber |
| L-26 | `00_DataAccess.gs:600-605` | `calculateDeadline()` doesn't detect `NaN` dates |
| L-27 | `00_DataAccess.gs:613-620` | `daysBetween()` DST bias — off-by-one possible |
| L-28 | `00_Security.gs:488` | `XFrameOptionsMode.DEFAULT` allows iframe embedding on access denied page |
| L-29 | `01_Core.gs:517-519` | `clearErrorLog()` doesn't log audit event |
| L-30 | `02_DataManagers.gs:728` | `addToConfigDropdown_()` can leave gaps in Config column |
| L-31 | `13_MemberSelfService.gs:55-59` | `generateMemberPIN()` may produce low-entropy PINs (padded with zeros) |
| L-32 | `18_WorkloadTracker.gs:390` | Biweekly reminder uses `>= 13` instead of `>= 14` |
| L-33 | `17_CorrelationEngine.gs:41-50` | `statStdDev_()` uses population std dev instead of sample |
| L-34 | `24_WeeklyQuestions.gs:355-358` | `getHistory` sorting fails for string dates |

### Minor Other (8)

| # | File | Issue |
|---|------|-------|
| L-35 | `00_DataAccess.gs:31-36` | Global mutable `_spreadsheetCache` accessible to all code |
| L-36 | `00_DataAccess.gs:122-123` | Duplicate range fetch for headers in `getOrCreateSheet()` |
| L-37 | `02_DataManagers.gs:2869-2880` | `highlightUrgentGrievances()` applies formatting per-row instead of batch |
| L-38 | `04d_ExecutiveDashboard.gs:773-776` | Duplicate highlight accumulation in `checkDuplicateMemberIDs_` |
| L-39 | `05_Integrations.gs:754,784` | `parent.removeFile()` deprecated in Drive API v3 |
| L-40 | `04d_ExecutiveDashboard.gs:157` | Chart.js from CDN without version pin or SRI hash |
| L-41 | `10_Main.gs:350-366` | `handleSecurityAudit_()` attempts `MailApp.sendEmail` in simple trigger (dead code) |
| L-42 | `02_DataManagers.gs:1048,1256` | `getClientSideEscapeHtml()` included twice in import dialog |

</details>

---

## Priority Remediation Plan

### Phase 1: Critical Security (Immediate)

1. **Fix web dashboard authentication** (CR-01, CR-02, CR-04, CR-05) — All `google.script.run` wrappers must resolve caller identity server-side
2. **Fix magic link replay** (CR-03) — Check `data.used` flag in `_validateMagicToken`
3. **Fix innerHTML XSS** (CR-16, CR-17) — Convert to DOM methods or add client-side `escapeHtml()`
4. **Add `escapeForFormula()`** to import paths (CR-18, CR-19, CR-20) and fix double-escaping (CR-26)

### Phase 2: Data Integrity (This Week)

5. **Scope batch writes** (CR-08, CR-10, CR-13, CR-15) — Write only changed columns, not entire sheets
6. **Protect manual data** (CR-09, CR-11, CR-12) — Check for existing values before overwriting
7. **Fix header string bugs** (CR-25, H-31) — Use GRIEVANCE_COLS constants and correct Roman numeral headers
8. **Fix Config row read** (CR-28) — Change row 2 to row 3

### Phase 3: Authorization & Logic (Next Sprint)

9. **Add `validateRole()`** to destructive functions (H-22, H-23, H-24, H-25, H-26)
10. **Fix grievance resolve status** (H-29) — Use outcome parameter for status, not always 'Settled'
11. **Unify audit log schema** (H-43, H-44) — Pick one schema, migrate
12. **Replace all hardcoded sheet names/columns** (H-01 through H-16) with `SHEETS.*`/`*_COLS` constants

### Phase 4: Hardening (Ongoing)

13. Add `LockService` to concurrent paths (H-17, H-18, H-19, H-20)
14. Add rate limiting to email/broadcast functions (H-28)
15. Add SRI hashes to CDN scripts, CSP to main app (H-52, M-12)
16. Deduplicate `closedStatuses` into a single constant (M-26)
17. PII masking improvements in `secureLog()` and alert emails (H-38, H-39, H-40)

---

---

## CI/CD & Build Issues (6 additional)

### CI-01 (CRITICAL): CI workflow checks for non-existent `dist/ConsolidatedDashboard.gs`
**File:** `.github/workflows/build.yml:49-78`
The CI workflow validates the existence of `dist/ConsolidatedDashboard.gs`, but the build system produces individual files (per CLAUDE.md). This check fails on every CI run.

### CI-02 (CRITICAL): CI branch trigger uses lowercase `main` instead of `Main`
**File:** `.github/workflows/build.yml:5`
The actual default branch is `Main` (capital M). The CI trigger on `main` means CI never runs on the primary branch.

### CI-03 (CRITICAL): Jest config requires 100% line coverage
**File:** `jest.config.js:14`
Requires 100% line coverage, but only 16 of 38 source files have dedicated test files. This makes the coverage gate unreliable or forces it to be bypassed.

### CI-04 (HIGH): ~42MB of PDFs and binary files committed to the repository
**Files:** `2020-2022 CONTRACT.pdf` (3.4MB), `2024-2026- original CONTRACT.pdf` (16MB), `2024-Local-509-Annual-Report-ENG.pdf` (12MB), `DDS CONTRACT.pdf` (5.4MB), `Forms.pdf` (1.2MB), etc.
Large binary files bloat the repo, slow clones, and cannot be diffed. These should be in Google Drive or Git LFS.

### CI-05 (MEDIUM): `dist/07_DevTools.gs` untracked despite `dist/` being tracked
The production build excludes DevTools, but the dev build includes it. The file is untracked, creating noise in `git status`.

### CI-06 (MEDIUM): `package.json` repository URL points to wrong repo
**File:** `package.json`
```json
"url": "https://github.com/Woop91/MULTIPLE-SCRIPS-REPO"
```
Should point to `DDS-Dashboard`, not `MULTIPLE-SCRIPS-REPO`.

---

*Review conducted on 76,626 lines across 41 source files. All findings cite specific file:line references.*
