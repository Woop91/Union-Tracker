# 🔍 DDS-Dashboard — Comprehensive Code Audit Report
# Date: 2026-03-10 | Version Audited: v4.25.7 | Auditor: Claude Opus 4.6

---

## 📊 EXECUTIVE SUMMARY

| Metric | Value |
|--------|-------|
| Source files (.gs) | 43 |
| HTML view files | 8 |
| Total source lines | 93,305 |
| Total test lines | 22,044 |
| Top-level functions | 1,380 |
| HTML→GS server calls | 117 verified |
| Orphaned/dead functions | 142 confirmed |
| Missing null guards | 32 |
| Hardcoded values (violations) | 7 |
| Data-write functions without locks | 5 |
| Files without test coverage | 14 |
| Deprecated functions still present | 30+ |
| HTML payload size | 965 KB |

### Verdict
The codebase is **architecturally sound** with strong patterns in authentication, dynamic configuration, and safe DOM building. However, there is significant technical debt from dead code (142 orphaned functions), missing null guards on sheet access (32 instances), and data-write functions without concurrency locks (5 instances). The HTML payload at 965 KB is a performance concern.

---

## 🟢 SECTION 1: WHAT'S WORKING WELL

### 1.1 Server Function Wiring — ✅ PERFECT
All 117 server functions called from HTML views (`serverCall().xxx()`) are properly defined in `.gs` files. Zero broken wires.

**Verification method:** Extracted all `serverCall()` chain endpoints from `steward_view.html`, `member_view.html`, `index.html`, `auth_view.html` and cross-referenced against 1,380 GS function definitions.

### 1.2 Authentication — ✅ COMPREHENSIVE
All 60 `data*` functions in `21_WebDashDataService.gs` have session token validation. Zero unauthenticated data endpoints.

### 1.3 No Duplicate Function Names — ✅ CLEAN
All 1,380 top-level functions have unique names across the entire GS codebase. Zero collisions in the GAS global namespace.

### 1.4 Build System — ✅ SOLID
- `build.js` includes syntax validation before copying
- src/dist parity is perfect (zero content differences)
- Production mode properly excludes `07_DevTools.gs`
- File count guard prevents exceeding GAS limits (52/60)
- Clean step prevents orphaned files in dist/

### 1.5 Sheet Name Constants — ✅ PROPERLY CENTRALIZED
`SHEETS` constant in `01_Core.gs` is the single source of truth. `SHEET_NAMES` is a clean alias (`var SHEET_NAMES = SHEETS`). Sheet names are used consistently via constants across 40+ files.

### 1.6 DOM Safety — ✅ STRONG
- `el()` function (1,796 calls across views) uses `textContent` and `createTextNode` — XSS-safe by design
- `html()` helper has explicit warning comment against unsanitized input
- `safeText()` wraps `escapeHtml()` for the rare `innerHTML` assignments with dynamic data
- Server-side `escapeHtml()`, `sanitizeObjectForHtml()`, `escapeForFormula()` all present

### 1.7 Script ID Security — ✅ CLEAN
DDS Apps Script ID (`18hHHX-...`) does not appear anywhere in the source files. Safe for Union-Tracker sync.

### 1.8 Config-Driven Architecture — ✅ MOSTLY DYNAMIC
Column identification by header name via `resolveColumnsFromSheet_()`. Sheet names from Config tab. Dropdowns from Config tab. Org name from Config tab (fixed in v4.20.13).

---

## 🔴 SECTION 2: CRITICAL ISSUES

### 2.1 Orphaned/Dead Code — 142 Functions (est. ~8,000+ lines)
**Impact:** Code bloat, confusion for maintainers, increased GAS payload size, potential for stale logic to be accidentally invoked.

These 142 functions are defined but never referenced in source, HTML, test files, or menu definitions:

**High-concern orphans (core modules):**
| Function | File | Lines | Concern |
|----------|------|-------|---------|
| `withErrorHandling` | 01_Core.gs:60 | ~17 | Error wrapper defined but unused everywhere |
| `mapMemberRow` | 01_Core.gs:2374 | ~42 | Row mapper defined but never called |
| `mapGrievanceRow` | 01_Core.gs:2416 | ~49 | Row mapper defined but never called |
| `runStartupValidation` | 01_Core.gs:366 | ~48 | Startup validator never triggered |
| `generateSequentialId` | 01_Core.gs:3252 | ~31 | ID generator never called |
| `getActionTypeConfig` | 01_Core.gs:2995 | ~234 | Massive config function unused |
| `calculateNextStepDeadline` | 02_DataManagers.gs:1859 | ~27 | Deadline calculator unused |
| `importMembersFromData` | 02_DataManagers.gs:1316 | ~108 | Import function unused |
| `validateWebAppRequest` | 00_Security.gs:392 | ~70 | Security validator unused |
| `getAccessDeniedPage` | 00_Security.gs:462 | ~33 | Access denied page unused |

**Deprecated but not removed (30+ functions):**
The codebase has 30+ `@deprecated` annotations but the functions remain. Includes the entire `04d_ExecutiveDashboard.gs` (1,010 lines, 23 functions) which is largely deprecated since v4.3.2.

**Recommendation:** Create a cleanup branch. Remove dead code in phases: (1) @deprecated functions first, (2) confirmed orphans by module, (3) verify with test suite between each phase.

### 2.2 Missing Null Guards on getSheetByName — 32 Instances
**Impact:** Runtime crash (`TypeError: Cannot call method getRange of null`) if a sheet is missing or renamed.

**Critical paths affected:**
```
02_DataManagers.gs:223  — generateMemberID_ (member creation breaks)
02_DataManagers.gs:2108 — recalcAllGrievancesBatched (batch recalc crashes)
02_DataManagers.gs:2250 — getGrievanceById (single lookup crashes)
02_DataManagers.gs:2273 — getOpenGrievances (dashboard data crashes)
03_UIComponents.gs:1379 — member stats collection
04c_InteractiveDashboard.gs:1514 — steward performance calc
05_Integrations.gs:63   — Drive integration init
05_Integrations.gs:3001 — web app dashboard stats
```

**Fix pattern:**
```javascript
// BEFORE (crashes if sheet missing)
var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
var data = sheet.getDataRange().getValues();

// AFTER (safe)
var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
if (!sheet) { throw new Error('Required sheet missing: ' + SHEETS.MEMBER_DIR); }
var data = sheet.getDataRange().getValues();
```

### 2.3 Data-Write Functions Without Lock Protection — 5 Functions
**Impact:** Concurrent users can cause race conditions, data corruption, duplicate rows, or lost writes.

| Function | File | Risk |
|----------|------|------|
| `syncMemberGrievanceData` | 02_DataManagers.gs:338 | Cross-sheet sync without lock |
| `updateMemberDataBatch` | 02_DataManagers.gs:951 | Batch member updates without lock |
| `addWebAppResource` | 05_Integrations.gs:4011 | Web resource creation without lock |
| `updateWebAppResource` | 05_Integrations.gs:4076 | Web resource update without lock |
| `deleteWebAppResource` | 05_Integrations.gs:4118 | Web resource deletion without lock |

**Fix:** Wrap in `withScriptLock_()` (already defined in `00_DataAccess.gs`).

---

## 🟡 SECTION 3: MODERATE ISSUES

### 3.1 Hardcoded Values (Design Principle Violations)

| File | Line | Hardcoded Value | Should Be |
|------|------|-----------------|-----------|
| `06_Maintenance.gs` | 3365 | `'_Archive_Grievances'` | `SHEETS.ARCHIVE_GRIEVANCES` or constant |
| `21_WebDashDataService.gs` | 568 | `'Grievance Log'` (fallback) | Remove fallback — SHEETS constant is canonical |
| `21_WebDashDataService.gs` | 575 | `'Member Directory'` (fallback) | Remove fallback |
| `21_WebDashDataService.gs` | 1525 | `'Contact Log'` | Should reference a constant |
| `21_WebDashDataService.gs` | 3000 | `'Contact Log'` | Same |
| `10b_SurveyDocSheets.gs` | 1992 | `'system@massability.org'` | `getConfigValue_(CONFIG_COLS.MAIN_CONTACT_EMAIL)` |
| `10b_SurveyDocSheets.gs` | 2007 | `'system@massability.org'` | Same |

### 3.2 Test Coverage Gaps — 14 Files Untested (14,853 lines)

| File | Lines | Concern |
|------|-------|---------|
| `08c_FormsAndNotifications.gs` | 2,444 | Forms, email, notifications — critical user-facing |
| `10a_SheetCreation.gs` | 2,261 | Sheet creation — infrastructure |
| `10b_SurveyDocSheets.gs` | 2,112 | Survey system — data integrity |
| `04c_InteractiveDashboard.gs` | 1,829 | Interactive dashboard — UI logic |
| `04b_AccessibilityFeatures.gs` | 1,030 | Accessibility — WCAG compliance |
| `04d_ExecutiveDashboard.gs` | 1,010 | Mostly deprecated but still present |
| `08a_SheetSetup.gs` | 959 | Sheet setup/validation |
| `04a_UIMenus.gs` | 929 | Menu definitions |
| `08b_SearchAndCharts.gs` | 884 | Search and chart generation |
| `08e_SurveyEngine.gs` | 848 | Survey processing engine |
| `08d_AuditAndFormulas.gs` | 2,179 | Audit logging, formulas |
| `10c_FormHandlers.gs` | 722 | Form submission handlers |
| `29_Migrations.gs` | 123 | Schema migration |
| `30_TestRunner.gs` | 1,654 | GAS-native test runner |

### 3.3 HTML Payload Size — 965 KB Total
**Impact:** Every page load transfers this entire payload. On slow connections, first meaningful paint could take 5-10+ seconds.

| File | Size | Concern |
|------|------|---------|
| `member_view.html` | 289 KB | Largest single view |
| `steward_view.html` | 229 KB | Second largest |
| `org_chart.html` | 217 KB | Third largest |
| `poms_reference.html` | 95 KB | Reference content |
| `styles.html` | 55 KB | Shared styles |
| `index.html` | 45 KB | SPA shell |

**Root cause:** All views are inlined into the GAS HTML output. No code splitting, lazy loading, or CDN for static assets.

### 3.4 Inconsistent var/const/let Usage
The codebase mixes `var` (traditional GAS) with `const`/`let` (V8 runtime). This is not a bug (GAS V8 supports both), but it indicates code written across different eras without standardization.

| File | const uses | let uses |
|------|-----------|----------|
| `12_Features.gs` | 380 | 12 |
| `02_DataManagers.gs` | 115 | 14 |
| `10_Main.gs` | 72 | 4 |
| `05_Integrations.gs` | 41 | 5 |

### 3.5 `withErrorHandling` Defined But Unused
The `withErrorHandling()` wrapper function at `01_Core.gs:60` was designed to wrap any function with standardized error handling, but is never called anywhere in the codebase. This means individual try-catch blocks are used everywhere instead of a centralized pattern — leading to inconsistent error handling.

### 3.6 Workload Tracker typeof Guards — Only 4 Found
Per SYNC-LOG.md, three files need `typeof` guards for Workload Tracker's absence in Union-Tracker. Only 4 guards found across the codebase — need to verify all necessary call sites are covered.

---

## 🔵 SECTION 4: INFORMATIONAL / LOW PRIORITY

### 4.1 Deprecated Code Still Shipped
`04d_ExecutiveDashboard.gs` (1,010 lines) is almost entirely `@deprecated` since v4.3.2 (replaced by modal dashboards). It still ships in production builds, adding ~44 KB to the GAS payload.

### 4.2 Event Bus Emitters Never Called
`emitDataChanged`, `emitEditEvent`, `emitFormEvent`, `emitSyncComplete` in `15_EventBus.gs` are all defined but never called. The Event Bus appears to be scaffolded but not integrated into the data mutation pipeline.

### 4.3 Hardcoded Org References in DevTools
`07_DevTools.gs` contains SEIU 509-specific emails (`bargaining@seiu509.org`, etc.) — acceptable since DevTools is excluded from production builds.

### 4.4 getLastRow() Without Empty-Sheet Safety
Several `getLastRow()` calls could return 0 on empty sheets, causing `getRange(2, 1, 0, ...)` which throws. Should use `Math.max(sheet.getLastRow() - 1, 0)`.

### 4.5 Constant Contact Integration Partially Orphaned
`checkConstantContactHealth()` is orphaned. The Constant Contact OAuth integration (`05_Integrations.gs:3099-3400`) uses hardcoded `https://localhost` as redirect URI — this only works in development.

### 4.6 Looker Integration Code — Large Unused Block
Functions `getLookerAnonStatus`, `getLookerConnectionUrl`, `getLookerStatus`, `setupLookerAnonIntegration`, `installLookerAllRefreshTrigger`, `installLookerRefreshTrigger`, `showLookerAnonConnectionHelp`, `showLookerConnectionHelp` in `12_Features.gs` — all orphaned. ~600+ lines of dead Looker integration code.

---

## 📋 SECTION 5: FEATURE TRACEABILITY MATRIX

### 5.1 Authentication Flow
```
User visits web app URL
  → doGet(e) [22_WebDashApp.gs]
    → doGetWebDashboard(e) [22_WebDashApp.gs]
      → SSO: getActiveUserEmail_() → builds SPA with inline session
      → Magic Link: authSendMagicLink() [19_WebDashAuth.gs] → email → token validation
      → PIN: verifyPIN() [13_MemberSelfService.gs]
    → Role check: getUserRole_() [00_Security.gs] → steward/member routing
    → SPA loads index.html → routes to steward_view.html or member_view.html
```
**Status:** ✅ Complete and wired

### 5.2 Grievance Lifecycle
```
Create: showNewGrievanceDialog() → startNewGrievance() → generateGrievanceId()
  → calculateInitialDeadlines() → writeToSheet → auditLog
Advance: advanceGrievanceStep() → getStepColumnSet() → calculateResponseDeadline()
  → updateRow → auditLog
Resolve: resolveGrievance() → updateStatus → calculateFinalDates → auditLog
View: dataGetStewardCases() → session validation → getSheetData → return to SPA
```
**Status:** ✅ Complete and wired

### 5.3 Member Management
```
Add: addMember() [withScriptLock_] → generateMemberID_ → writeToSheet → auditLog
Update: updateMember() [withScriptLock_] → findRow → updateCells → auditLog
Import: showImportMembersDialog() → importMembersBatch() [withScriptLock_] → batch write
Search: searchMembers() → regex match → return results
Profile: dataGetFullProfile() → session check → aggregate member + grievance data
```
**Status:** ✅ Complete and wired (import path has an unused `importMembersFromData` function)

### 5.4 Survey System
```
Config: SURVEY_QUESTIONS sheet → dataGetSurveyQuestions()
Submit: dataSubmitSurveyResponse() → validate → hash email → write to vault
  → update tracking → anonymize for satisfaction sheet
View: dataGetSatisfactionSummary() → aggregate from satisfaction sheet
Admin: dataOpenNewSurveyPeriod() → reset tracking → notify members
```
**Status:** ✅ Wired (some form trigger wiring unconfirmed per AI_REFERENCE.md)

### 5.5 Notification System
```
Send: sendWebAppNotification() → write to Notifications sheet
Fetch: getWebAppNotifications() → filter by recipient → return
Dismiss: dismissWebAppNotification() → update status column
Archive: archiveWebAppNotification() → move to archive
Count: getWebAppNotificationCount() → count unread
```
**Status:** ✅ Complete and wired

### 5.6 Q&A Forum
```
Ask: qaSubmitQuestion() → validate → write to QA_FORUM sheet → notify stewards
Answer: qaSubmitAnswer() → validate → write to QA_ANSWERS sheet → notify asker
Moderate: qaModerateQuestion() / qaModerateAnswer() → steward-only → flag/approve
Resolve: qaResolveQuestion() → mark resolved → archive
View: qaGetQuestions() → filter by status → return with answers
```
**Status:** ✅ Complete and wired

### 5.7 Weekly Questions / Pulse Surveys
```
Set: wqSetStewardQuestion() → write to WEEKLY_QUESTIONS sheet
Draw: wqManualDrawCommunityPoll() → pick from QUESTION_POOL
Submit: wqSubmitResponse() → hash email → write to WEEKLY_RESPONSES
View: wqGetActiveQuestions() → filter active → return
Close: wqClosePoll() → update status → aggregate results
```
**Status:** ✅ Complete and wired

### 5.8 Timeline / Activity Feed
```
Add: tlAddTimelineEvent() → validate → write to TIMELINE_EVENTS sheet
View: tlGetTimelineEvents() → paginate → return
Delete: tlDeleteTimelineEvent() → steward-only → soft delete
Import: tlImportCalendarEvents() → fetch from Google Calendar → write
```
**Status:** ✅ Wired (tlUpdateTimelineEvent and tlAttachDriveFiles defined but orphaned — possibly incomplete features)

### 5.9 Workload Tracker
```
Submit: processWorkloadFormSSO() → validate → write to WORKLOAD_VAULT
Dashboard: getWorkloadDashboardDataSSO() → aggregate from vault + reporting
History: getWorkloadHistorySSO() → filter by user → return
Reminder: setWorkloadReminderSSO() / getWorkloadReminderSSO() → manage prefs
Export: exportWorkloadHistoryCSV() → build CSV → return download
```
**Status:** ✅ Complete and wired (excluded from Union-Tracker per SYNC-LOG.md)

### 5.10 Failsafe / Data Protection
```
Setup: fsSetupTriggers() → install backup/health triggers
Backup: fsBackupCriticalSheets() → copy critical sheets to backup folder
Diagnostic: fsDiagnostic() → check sheet health → return report
Digest: fsGetDigestConfig() / fsUpdateDigestConfig() → manage email digest prefs
Export: fsTriggerBulkExport() → export all data sheets
```
**Status:** ✅ Complete and wired

---

## 📋 SECTION 6: BUTTON/INTERACTION AUDIT

### 6.1 Steward View — Buttons & Actions
All buttons in `steward_view.html` use the `el()` DOM builder with `onclick` event handlers that call `serverCall()` chains. Verified handlers:

| UI Element | Handler Chain | Status |
|-----------|--------------|--------|
| Logout button | `handleLogout()` → `authLogout()` | ✅ |
| View member profile | `serverCall().dataGetFullProfile()` | ✅ |
| Log contact | `serverCall().dataLogMemberContact()` | ✅ |
| Create task | `serverCall().dataCreateTask()` | ✅ |
| Complete task | `serverCall().dataCompleteTask()` | ✅ |
| Send broadcast | `serverCall().dataSendBroadcast()` | ✅ |
| Send direct message | `serverCall().dataSendDirectMessage()` | ✅ |
| View case folder | `serverCall().dataGetMemberCaseFolderUrl()` | ✅ |
| Submit Q&A answer | `serverCall().qaSubmitAnswer()` | ✅ |
| Resolve question | `serverCall().qaResolveQuestion()` | ✅ |
| Add timeline event | `serverCall().tlAddTimelineEvent()` | ✅ |
| Survey management | `serverCall().dataOpenNewSurveyPeriod()` | ✅ |
| Weekly questions | `serverCall().wqSetStewardQuestion()` | ✅ |
| Resource management | `serverCall().addWebAppResource()` | ✅ |
| Test runner | `serverCall().dataRunTests()` | ✅ |
| Failsafe controls | `serverCall().fsSetupTriggers()` | ✅ |

### 6.2 Member View — Buttons & Actions
| UI Element | Handler Chain | Status |
|-----------|--------------|--------|
| Logout button | `handleLogout()` → `authLogout()` | ✅ |
| View grievances | `serverCall().dataGetMemberGrievances()` | ✅ |
| Start grievance draft | `serverCall().dataStartGrievanceDraft()` | ✅ |
| Request case folder | `serverCall().dataCreateGrievanceDrive()` | ✅ |
| Choose steward | `serverCall().dataAssignSteward()` | ✅ |
| Update profile | `serverCall().dataUpdateProfile()` | ✅ |
| Submit feedback | `serverCall().dataSubmitFeedback()` | ✅ |
| Q&A submit question | `serverCall().qaSubmitQuestion()` | ✅ |
| Q&A upvote | `serverCall().qaUpvoteQuestion()` | ✅ |
| Submit survey | `serverCall().dataSubmitSurveyResponse()` | ✅ |
| Weekly question response | `serverCall().wqSubmitResponse()` | ✅ |
| View resources | `serverCall().getWebAppResourceLinks()` | ✅ |

---

## 📋 SECTION 7: RECOMMENDATIONS (PRIORITIZED)

### P0 — Fix Before Next Deploy
1. **Add null guards** to all 32 `getSheetByName` calls without them
2. **Add `withScriptLock_`** to the 5 unprotected data-write functions
3. **Replace hardcoded `system@massability.org`** with config value (2 locations)
4. **Replace hardcoded `'_Archive_Grievances'`** with a SHEETS constant

### P1 — Fix This Sprint
5. **Remove `'Grievance Log'` / `'Member Directory'` fallback strings** in `21_WebDashDataService.gs` (4 locations) — these are from before the SHEETS constant and mask bugs
6. **Remove `'Contact Log'` hardcoded string** (2 locations same file)
7. **Add tests** for the 5 most critical untested files: `08c_FormsAndNotifications.gs`, `10a_SheetCreation.gs`, `10b_SurveyDocSheets.gs`, `08d_AuditAndFormulas.gs`, `04c_InteractiveDashboard.gs`

### P2 — Fix This Month
8. **Dead code cleanup** — remove 142 orphaned functions in phases:
   - Phase 1: Remove `@deprecated` functions (30+)
   - Phase 2: Remove orphaned Looker integration (~600 lines)
   - Phase 3: Remove remaining orphans by module
9. **Integrate `withErrorHandling()`** or remove it — currently defined but unused
10. **Wire up Event Bus** or remove `emitDataChanged`, `emitEditEvent`, `emitFormEvent`, `emitSyncComplete`
11. **Standardize var/const/let** — adopt `const` by default, `let` for reassignment

### P3 — Backlog
12. **HTML payload optimization** — investigate lazy loading views, code splitting, or CDN for Chart.js
13. **Remove deprecated `04d_ExecutiveDashboard.gs`** entirely (1,010 lines)
14. **Add `getLastRow()` safety** (`Math.max(sheet.getLastRow() - 1, 0)`) in edge cases
15. **Complete timeline feature** — `tlUpdateTimelineEvent` and `tlAttachDriveFiles` are defined but orphaned
16. **Confirm survey form trigger wiring** — `onSatisfactionFormSubmit` trigger may not be installed

---

## ❓ UNRESOLVED QUESTIONS

1. `syncMemberGrievanceData()` — intentionally lockless or oversight?
2. Web app resource CRUD (add/update/delete) — are these safe without locks given single-steward usage?
3. EventBus (15_EventBus.gs) — planned for future integration or abandoned feature?
4. Looker integration (12_Features.gs ~3066-3999) — planned or permanently abandoned?
5. `tlUpdateTimelineEvent` / `tlAttachDriveFiles` — stub for future or forgotten?
6. `importMembersFromData` vs `importMembersBatch` — are both needed or is one dead?
7. `04d_ExecutiveDashboard.gs` — safe to fully remove or does any menu still reference it?
8. Survey form trigger (`onSatisfactionFormSubmit`) — currently installed on live sheet?
9. Constant Contact OAuth with `https://localhost` redirect — is this integration active?
10. `_Archive_Grievances` — does a SHEETS constant exist for this? If not, should one be added?
