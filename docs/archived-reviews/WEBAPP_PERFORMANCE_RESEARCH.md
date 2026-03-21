# Web App Performance Research

**Date:** 2026-03-12
**Scope:** SolidBase web app (GAS-hosted SPA)
**Goal:** Identify all opportunities to improve load time, responsiveness, and perceived performance.

---

## Table of Contents

1. [Current Architecture Summary](#1-current-architecture-summary)
2. [Existing Optimizations](#2-existing-optimizations-already-in-place)
3. [Cache Warming — Can It Be Done?](#3-cache-warming--can-it-be-done)
4. [Recommendations — Server-Side](#4-recommendations--server-side)
5. [Recommendations — Client-Side](#5-recommendations--client-side)
6. [Recommendations — Network & Loading](#6-recommendations--network--loading)
7. [Recommendations — Architectural](#7-recommendations--architectural)
8. [Priority Matrix](#8-priority-matrix)
9. [Implementation Notes](#9-implementation-notes)

---

## 1. Current Architecture Summary

The web app is a GAS-hosted SPA served via `doGet()`. The critical path is:

```
User visits URL
  → doGet() (22_WebDashApp.gs)
    → ConfigReader.getConfig()          [CacheService, 6hr TTL]
    → Auth.resolveUser(e)               [SSO / magic link / session token]
    → DataService.findUserByEmail()     [_getCachedSheetData, 2min TTL]
    → _serveDashboard() → HtmlService.createTemplateFromFile('index')
      → Injects PAGE_DATA as JSON
      → Evaluates template (includes styles.html, steward_view.html, member_view.html)
  → Browser receives single HTML document (~200-300KB)
    → Parses inline CSS + JS
    → Loads Google Fonts (4 families)
    → Calls initApp()
      → google.script.run.dataGetBatchData(SESSION_TOKEN)
        → Server: resolves auth, fetches all initial data in one call
      → Renders dashboard with batch data
```

**Key constraint:** Every `google.script.run` call is a full HTTP round-trip through GAS infrastructure (~300-800ms per call). Minimizing these calls is the single highest-impact optimization.

---

## 2. Existing Optimizations Already in Place

| Optimization | Location | Notes |
|---|---|---|
| **Batch data pre-fetch** | `dataGetBatchData()` in 21_WebDashDataService.gs:3237 | Single round-trip loads all initial view data |
| **Two-tier server cache** | `_getCachedSheetData()` in 21_WebDashDataService.gs:2082 | In-memory + CacheService (2min TTL) |
| **Config cache** | ConfigReader in 20_WebDashConfigReader.gs:16 | CacheService with 6hr TTL |
| **Column map persistence** | `RESOLVED_COL_MAPS` in 01_Core.gs | 6hr CacheService TTL avoids re-scanning headers |
| **Client DataCache** | index.html:302-346 | In-memory cache with 2min default TTL |
| **localStorage session** | index.html:32-62 | Persists session token across reloads |
| **Chart.js lazy load** | ChartHelper in index.html:890 | Only fetched when first chart rendered |
| **Lazy-loaded views** | Org Chart, POMS Reference | Fetched only on tab click |
| **Debouncing** | Search (200ms), badges (600ms), resize (200ms) | Prevents redundant work |
| **`display=swap` fonts** | index.html:11 | FOUT instead of FOIT |
| **Preconnect** | index.html:9-10 | DNS prefetch for Google Fonts |
| **Reduced motion** | styles.html:1920 | Disables animations for a11y + perf |
| **Calendar event cache** | 21_WebDashDataService.gs:2036 | 15min TTL on events |
| **Survey summary cache** | 08e_SurveyEngine.gs | 10min TTL |
| **Rate limiting via cache** | Email (60s), Magic Link (15min), Workload (1hr) | CacheService-based |

---

## 3. Cache Warming — Can It Be Done?

### Yes, and partially already implemented

**Existing:** `warmUpCaches()` in `06_Maintenance.gs:753` pre-loads grievances, members, stewards, and dashboard metrics into the two-tier cache (CacheService + PropertiesService). It's available via the admin menu: **Data Integrity > Warm Up Caches**.

### What it currently warms

| Cache Key | Loader | TTL |
|---|---|---|
| `cache_grievances` | Full Grievance Log sheet | 5min (MEMORY_TTL) |
| `cache_members` | Full Member Directory sheet | 5min |
| `cache_stewards` | Filtered steward list | 5min |
| `cache_metrics` | Pre-computed KPIs | 5min |

### What it does NOT warm (gaps)

| Cache Key | What | Current TTL | Impact |
|---|---|---|---|
| `SD_Grievance_Log` | Web app sheet data cache | 2min | First web app load after cache expiry reads sheet fresh |
| `SD_Member_Directory` | Web app sheet data cache | 2min | Same |
| `ORG_CONFIG` | Org configuration | 6hr | Low impact (long TTL) |
| `events_<calendarId>` | Upcoming events | 15min | Calendar widget cold on first load |
| `satisfactionSummary_*` | Survey aggregates | 10min | Insights tab cold start |

### Recommended: Automated cache warming via time-based trigger

```
Approach: Install a time-based trigger (every 5 minutes) that calls a
lightweight warmWebAppCaches_() function. This keeps the web app's
CacheService entries hot so users never hit a cold cache.
```

**Why this matters:** The web app's `_getCachedSheetData()` (2min TTL) and the spreadsheet-side `getCachedData()` (5min TTL) are **separate cache namespaces** with different key prefixes (`SD_*` vs `cache_*`). Warming one doesn't warm the other. A dedicated web-app-aware warmer would populate `SD_*` keys that `dataGetBatchData()` actually reads.

### Proposed: `warmWebAppCaches_()`

```
1. Read Grievance Log → cache as SD_Grievance_Log (2min)
2. Read Member Directory → cache as SD_Member_Directory (2min)
3. Read Steward Tasks → cache as SD_Steward_Tasks (2min)
4. Read Config → cache as ORG_CONFIG (already 6hr, just ensure hot)
5. Read Calendar events → cache as events_<id> (15min)
```

**Trigger frequency:** Every 2 minutes (matches TTL) would be ideal but costs GAS quota. Every 5 minutes is a practical compromise — users may see one cold load per 5min window.

**GAS Quota consideration:** Each warm-up reads 3-5 sheets (~3-5 seconds of execution). At every 5 minutes = 288 executions/day = ~15 minutes of total execution time. Well within the 90min/day limit.

---

## 4. Recommendations — Server-Side

### S1. Route `getAllMembers()` through `_getCachedSheetData()` [HIGH IMPACT]

**Problem:** `getAllMembers()` (21_WebDashDataService.gs:857) calls `sheet.getDataRange().getValues()` directly, bypassing the `_getCachedSheetData()` two-tier cache. It's called from 5+ locations including batch data.

**Fix:** Refactor to use cached sheet data:
```js
function getAllMembers() {
  var cached = _getCachedSheetData(MEMBER_SHEET);
  if (!cached || !cached.data) return [];
  // ... build member list from cached.data instead of direct read
}
```

**Expected savings:** Eliminates 1-5 redundant full sheet reads per batch data call.

---

### S2. Batch the badge refresh into a single server function [HIGH IMPACT]

**Problem:** `_refreshNavBadges()` (steward_view.html:429-481) makes 3 serial `google.script.run` calls:
1. `dataGetStewardKPIs()`
2. `dataGetTasks()`
3. `qaGetQuestions()`

Each call is 300-800ms. Total: 900-2400ms of serial latency.

**Fix:** Create `dataGetBadgeCounts(sessionToken)` that returns all badge data in one call:
```js
{ kpis: {...}, taskCount: N, overdueTaskCount: N, qaUnansweredCount: N }
```

**Expected savings:** 600-1600ms per badge refresh (2 round-trips eliminated).

---

### S3. Add in-execution cache for `ConfigReader.getConfig()` [MEDIUM]

**Problem:** `ConfigReader.getConfig()` is called from multiple server functions within the same request. While CacheService has a 6hr TTL, each `cache.get()` + `JSON.parse()` still costs ~5-10ms.

**Fix:** Add a module-scoped variable:
```js
var _configMemo = null;
function getConfig(forceRefresh) {
  if (!forceRefresh && _configMemo) return _configMemo;
  // ... existing logic ...
  _configMemo = config;
  return config;
}
```

**Expected savings:** ~20-50ms per request with multiple config reads.

---

### S4. Consolidate ConfigReader cell reads into a single row read [LOW]

**Problem:** `_readCell()` (20_WebDashConfigReader.gs:133-140) reads 12+ individual cells one at a time.

**Fix:** Read entire row 3 in one call:
```js
var row = sheet.getRange(3, 1, 1, sheet.getLastColumn()).getValues()[0];
```

**Expected savings:** Minimal (mitigated by 6hr cache), but reduces cold-start config read from ~12 API calls to 1.

---

### S5. Cache `_findUserByName()` results [MEDIUM]

**Problem:** `_findUserByName()` (21_WebDashDataService.gs:1267) performs a direct `getDataRange().getValues()` read, not routed through `_getCachedSheetData()`.

**Fix:** Route through cached sheet data, similar to S1.

---

### S6. Pre-compute and cache steward KPIs alongside cases [MEDIUM]

**Problem:** `_computeKPIsFromCases()` is called after fetching cases, re-iterating the full case array. If cases are cached, KPIs should be cached alongside them.

**Fix:** Store computed KPIs in the same cache entry as cases, or in a parallel cache key.

---

## 5. Recommendations — Client-Side

### C1. Render initial view from PAGE_DATA without waiting for batch [HIGH IMPACT]

**Problem:** `initApp()` waits for `dataGetBatchData()` to complete before rendering anything. Users see a blank/loading screen for 500-2000ms.

**Fix:** Render a skeleton UI immediately from PAGE_DATA (user name, role, nav structure, empty card placeholders), then hydrate with batch data when it arrives.

```js
function initApp() {
  // Render skeleton immediately
  renderSkeleton(document.getElementById('app'), CURRENT_VIEW);

  // Fetch batch data in background
  google.script.run
    .withSuccessHandler(function(batch) {
      AppState.batchData = batch;
      hydrateDashboard(batch);
    })
    .dataGetBatchData(SESSION_TOKEN);
}
```

**Expected improvement:** Perceived load time drops from ~1-2s to <200ms (skeleton appears instantly).

---

### C2. Implement pagination or virtual scrolling for case/member lists [MEDIUM]

**Problem:** All cases and members are rendered to DOM at once. No pagination observed. With >200 items, DOM creation becomes noticeable.

**Fix options:**
- **Simple pagination:** Render first 50 items, add "Load More" button
- **Intersection Observer:** Load next batch when user scrolls near bottom
- **Virtual scrolling:** Only render visible items (more complex, higher payoff for 500+ items)

**Recommended:** Intersection Observer approach — minimal code, works well for expected dataset sizes (<500).

---

### C3. Cache batch data in localStorage with ETag pattern [MEDIUM]

**Problem:** `DataCache` (in-memory) is lost on page reload. Every page load requires a fresh `dataGetBatchData()` round-trip.

**Fix:** Store batch data in localStorage with a timestamp. On load:
1. Immediately render from localStorage (stale data)
2. Fetch fresh data in background
3. Update UI only if data changed

```js
// On page load:
var stale = JSON.parse(localStorage.getItem('batchData') || 'null');
if (stale && (Date.now() - stale.ts) < 120000) {
  AppState.batchData = stale.data;
  renderDashboard();  // Instant render from stale data
}
// Always fetch fresh in background
google.script.run.withSuccessHandler(function(fresh) {
  localStorage.setItem('batchData', JSON.stringify({ data: fresh, ts: Date.now() }));
  if (JSON.stringify(fresh) !== JSON.stringify(stale?.data)) {
    AppState.batchData = fresh;
    renderDashboard();  // Update only if changed
  }
}).dataGetBatchData(SESSION_TOKEN);
```

**Expected improvement:** Returning users see dashboard instantly (0ms perceived load).

---

### C4. Parallelize badge refresh calls [LOW-MEDIUM]

**Problem:** Badge refresh uses nested callbacks (serial). Even if we don't combine them server-side (S2), they can run in parallel client-side.

**Fix:** Fire all 3 calls simultaneously, track completion with a counter:
```js
var pending = 3;
function done() { if (--pending === 0) rerenderNav(); }

serverCall().withSuccessHandler(function(kpis) { AppState.kpis = kpis; done(); }).dataGetStewardKPIs(SESSION_TOKEN);
serverCall().withSuccessHandler(function(tasks) { /* process */ done(); }).dataGetTasks(SESSION_TOKEN, null);
serverCall().withSuccessHandler(function(qa) { /* process */ done(); }).qaGetQuestions(SESSION_TOKEN, 1, 999, 'recent');
```

**Expected savings:** Reduces badge refresh from 3x serial (900-2400ms) to 1x parallel (300-800ms).

---

### C5. Use `requestIdleCallback` for non-critical card rendering [LOW]

**Problem:** All cards render synchronously in a `forEach` loop. Large lists block the main thread.

**Fix:** Render first 10 cards immediately, remaining in idle callbacks:
```js
var INITIAL_BATCH = 10;
items.slice(0, INITIAL_BATCH).forEach(renderCard);
var remaining = items.slice(INITIAL_BATCH);
function renderNext() {
  if (!remaining.length) return;
  var batch = remaining.splice(0, 10);
  batch.forEach(renderCard);
  requestIdleCallback(renderNext);
}
requestIdleCallback(renderNext);
```

---

### C6. Increase DataCache TTL for stable data [LOW]

**Problem:** DataCache default TTL is 2 minutes. Some data (member directory, config) changes infrequently.

**Fix:** Use longer TTLs for stable data:
- Member directory: 5 minutes
- Steward directory: 5 minutes
- Config data: 30 minutes
- Case list: 2 minutes (keep current — changes frequently)
- KPIs: 2 minutes

---

## 6. Recommendations — Network & Loading

### N1. Reduce Google Fonts payload [MEDIUM]

**Problem:** Loading 4 font families (JetBrains Mono, Space Grotesk, Plus Jakarta Sans, Sora) with multiple weights. This is ~150-200KB of font data.

**Fix options:**
- **Drop one display font:** Space Grotesk and Sora serve similar purposes. Pick one.
- **Subset fonts:** Add `&text=` parameter for JetBrains Mono (only used for code) to load Latin subset only
- **Use `font-display: optional`** instead of `swap` for non-critical fonts (prevents layout shift)
- **Self-host fonts** as base64 in styles.html (eliminates DNS lookup + CDN round-trip, but increases HTML size)

**Expected savings:** 50-100KB reduction in font payload, 100-300ms faster first paint.

---

### N2. Defer non-critical CSS [LOW]

**Problem:** All ~1,900 lines of CSS are loaded inline before any content renders. Nav theme animations, org chart styles, etc. aren't needed on initial paint.

**Fix:** Split critical CSS (layout, colors, typography) from non-critical (animations, themes, specific views). Inline critical CSS in `<head>`, defer the rest.

**Feasibility:** Low — GAS templating makes CSS splitting awkward. Would require build-time CSS extraction.

---

### N3. Preload the batch data call [MEDIUM]

**Problem:** The batch data `google.script.run` call only fires after all JS has parsed and `initApp()` executes.

**Fix:** Move the `google.script.run.dataGetBatchData()` call to the very top of the `<script>` block, before ThemeEngine and other module definitions:
```html
<script>
  var PAGE_DATA = <?!= pageData ?>;
  var _batchPromise = null;
  if (PAGE_DATA.view === 'steward' || PAGE_DATA.view === 'member') {
    var _batchResolve;
    _batchPromise = { data: null, ready: false, callbacks: [] };
    google.script.run
      .withSuccessHandler(function(batch) {
        _batchPromise.data = batch;
        _batchPromise.ready = true;
        _batchPromise.callbacks.forEach(function(cb) { cb(batch); });
      })
      .dataGetBatchData(PAGE_DATA.sessionToken || '');
  }
  // ... rest of JS modules parse while batch is in-flight ...
</script>
```

**Expected savings:** 100-300ms (batch call starts during JS parse instead of after).

---

## 7. Recommendations — Architectural

### A1. Implement time-based cache warming trigger [HIGH IMPACT]

As described in Section 3. Install a trigger that runs every 5 minutes to keep `SD_*` cache keys warm. This ensures the first user to load the app within any 5-minute window doesn't pay the cold-cache penalty.

**Implementation:**
```js
function setupCacheWarmingTrigger() {
  // Remove existing to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'warmWebAppCaches_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('warmWebAppCaches_')
    .timeBased()
    .everyMinutes(5)
    .create();
}

function warmWebAppCaches_() {
  var cache = CacheService.getScriptCache();
  // Warm the 3 most-read sheets
  var sheets = [SHEETS.GRIEVANCE_LOG, SHEETS.MEMBER_DIRECTORY, SHEETS.STEWARD_TASKS];
  sheets.forEach(function(name) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    // Serialize with date handling
    var serializable = data.map(function(row) {
      return row.map(function(cell) {
        return cell instanceof Date ? { __d: cell.toISOString() } : cell;
      });
    });
    var json = JSON.stringify({ data: serializable });
    var key = 'SD_' + name.replace(/\s/g, '_');
    if (json.length < 95000) {
      cache.put(key, json, 300); // 5min TTL to bridge between trigger runs
    }
  });
  // Also warm config
  if (typeof ConfigReader !== 'undefined') ConfigReader.getConfig(true);
}
```

---

### A2. Add a `dataGetBadgeCounts()` endpoint [HIGH IMPACT]

Consolidate the 3 serial badge refresh calls into one. See S2 above.

---

### A3. Consider HTML template pre-evaluation [LOW-MEDIUM]

**Problem:** `HtmlService.createTemplateFromFile('index').evaluate()` processes scriptlets and includes on every `doGet()` call. The template content rarely changes.

**Observation:** GAS doesn't support caching evaluated HTML. But the `include()` calls for `styles.html`, `steward_view.html`, `member_view.html` could be pre-concatenated at build time into a single HTML file, reducing template evaluation overhead.

**Feasibility:** Would require build script changes. Moderate effort.

---

### A4. Add LockService for concurrent write protection [LOW — CORRECTNESS]

**Problem:** No LockService usage detected. Concurrent writes to the same sheet row (e.g., two stewards updating the same grievance) could cause data races.

**Fix:** Add `LockService.getScriptLock()` around write operations. Not strictly a performance improvement, but prevents data corruption that could cause cascading re-reads.

---

## 8. Priority Matrix

| ID | Recommendation | Impact | Effort | Priority |
|---|---|---|---|---|
| **C1** | Skeleton UI before batch data | Very High | Low | **P0** |
| **N3** | Preload batch data call | High | Low | **P0** |
| **S2/A2** | Batch badge refresh endpoint | High | Low | **P1** |
| **S1** | Route getAllMembers() through cache | High | Low | **P1** |
| **A1** | Time-based cache warming trigger | High | Medium | **P1** |
| **C3** | localStorage stale-while-revalidate | High | Medium | **P1** |
| **C4** | Parallelize badge refresh calls | Medium | Low | **P2** |
| **S3** | In-execution config memo | Medium | Low | **P2** |
| **N1** | Reduce Google Fonts payload | Medium | Low | **P2** |
| **C2** | Pagination / virtual scrolling | Medium | Medium | **P2** |
| **S5** | Cache _findUserByName() | Medium | Low | **P2** |
| **S6** | Pre-compute KPIs with cases | Medium | Low | **P2** |
| **C5** | requestIdleCallback for cards | Low | Low | **P3** |
| **C6** | Increase DataCache TTLs | Low | Low | **P3** |
| **S4** | Single-row config read | Low | Low | **P3** |
| **N2** | Defer non-critical CSS | Low | High | **P3** |
| **A3** | Pre-evaluate HTML templates | Low-Med | Medium | **P3** |
| **A4** | LockService for writes | Low (perf) | Medium | **P3** |

---

## 9. Implementation Notes

### Quick Wins (can ship today)
1. **C1 + N3:** Skeleton UI + preloaded batch call — pure frontend changes in `index.html`
2. **C4:** Parallelize badge calls — change `_refreshNavBadges()` in `steward_view.html`
3. **S1:** Route `getAllMembers()` through `_getCachedSheetData()` — small refactor in `21_WebDashDataService.gs`
4. **S3:** Config memo — 3 lines added to `20_WebDashConfigReader.gs`

### Medium Effort (1-2 sessions)
5. **S2/A2:** New `dataGetBadgeCounts()` endpoint + client update
6. **A1:** Cache warming trigger + `warmWebAppCaches_()` function
7. **C3:** localStorage stale-while-revalidate pattern

### GAS-Specific Constraints
- **CacheService limit:** 100KB per key. Large sheets (>1000 rows) may exceed this. The existing 95KB guard handles this gracefully by falling back to direct reads.
- **Trigger quota:** Max 20 triggers per script. Cache warming trigger adds 1.
- **Execution time:** 6-minute max per execution. Cache warming should complete in <10 seconds.
- **No HTTP caching headers:** GAS web apps don't support custom `Cache-Control` headers. All caching must be done via CacheService (server) or JavaScript (client).
- **No service workers:** GAS iframe sandbox prevents service worker registration.
- **No WebSockets/SSE:** GAS doesn't support persistent connections. All real-time patterns must use polling or event-driven refresh.

### Estimated Combined Impact

If P0 + P1 items are implemented:
- **First load (cold):** ~2-3s → ~1-1.5s (skeleton + preloaded batch + warm cache)
- **Return visit:** ~1-2s → ~0-200ms (localStorage stale-while-revalidate)
- **Badge refresh:** ~1-2.5s → ~0.3-0.8s (single endpoint)
- **Tab switching:** Already fast (~50ms from DataCache) — no change needed
