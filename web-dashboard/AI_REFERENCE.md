# AI REFERENCE DOCUMENT — Grievance Dashboard Web App
# ⚠️ THIS FILE MUST NEVER BE DELETED. ONLY APPEND. ⚠️
# Used by: Claude, Gemini, ChatGPT, or any LLM working on this codebase.
# Last updated: 2026-02-21

---

## 🏗️ PROJECT OVERVIEW

**What:** Mobile-optimized web app dashboard for workplace organization grievance tracking & member management.
**Served by:** Google Apps Script `doGet()` → HtmlService, bound to a Google Sheet.
**Target users:** Union stewards (power users) and members (casual users).
**Repo:** `Woop91/DDS-Dashboard` → branch: `web-dashboard`
**All other repos are DEPRECATED** (509-dashboard, union-steward-app — do not reference).

---

## 🔴 CRITICAL RULES — READ FIRST

1. **EVERYTHING MUST BE DYNAMIC.** Never hardcode org names, unit names, column positions, sheet names, or any data that lives in the spreadsheet. Always read from the sheet or the Config tab.
2. **Config tab is the single source of truth** for org-specific settings (name, abbreviation, accent color, units, magic link expiry, cookie duration).
3. **Role column exists** in Member Directory. Read it dynamically — do not assume column position. Find by header name.
4. **Bound to existing sheet** — the web app code lives in the same Apps Script project as the existing 15k-line dashboard code. Do NOT create a separate script project.
5. **Bitly redirect** — users arrive via a Bitly short link. The raw Apps Script deployment URL changes on each deploy. Bitly is updated manually. The app should never expose or reference the raw script.google.com URL to users.
6. **One repo only** — `Woop91/DDS-Dashboard`, branch `web-dashboard`. All pushes go here.
7. **Auth: SSO + Magic Link** — both options available to all users. Role determined by directory lookup post-auth, not by auth method.
8. **Dual-role users** default to steward view with a toggle to switch to member view.
9. **No member PII** visible to other members. Members see only their own grievances.
10. **Steward contact** — email always shown, phone only if steward has one listed.

---

## 🎨 DESIGN SYSTEM

### Steward View — "Mono Signal"
- Dark default, light mode available
- Font body: JetBrains Mono (monospace)
- Font display: Space Grotesk
- Card radius: 10px
- Status signaling: colored dot with glow
- Layout: dense, filterable case list, 4-up KPI bar
- Features: bulk select, flag system, FAB for new case, notifications

### Member View — "Glass Depth"
- Dark default, light mode available
- Font body: Plus Jakarta Sans
- Font display: Sora
- Card radius: 20px
- Status signaling: bottom glow bar on cards
- Layout: single grievance card with countdown, timeline, resources
- Features: steward contact, "Know Your Rights", FAQ, file grievance

### Theme Engine
- Accent color set per org in Config tab (stored as hue value 0-360)
- Full palette generated dynamically from accent hue
- Both roles share the same accent hue but apply it differently

### Auth Screen
- Single landing page for all users
- "Continue with Google" (SSO) + "Sign in with Email" (magic link)
- Magic link: 7-day expiry
- Remember device: 30-day cookie
- Role determined post-auth by directory lookup

---

## 📁 FILE STRUCTURE

```
/WebApp.gs        — doGet() handler, routing, page serving
/Auth.gs          — SSO detection, magic link generation/validation, cookie management
/ConfigReader.gs  — reads Config tab, returns org settings
/DataService.gs   — reads Member Directory & Grievance Log, returns JSON for frontend
/index.html       — SPA shell, loads CSS + JS, handles client-side routing
/auth.html        — login screen (SSO + magic link)
/steward.html     — steward dashboard view (Mono Signal)
/member.html      — member dashboard view (Glass Depth)
/styles.html      — shared CSS (theme engine, components)
/app.js.html      — client-side JavaScript (bundled in <script> tags for HtmlService)
/AI_REFERENCE.md  — THIS FILE
```

Note: Apps Script serves HTML files via HtmlService. JS and CSS must be embedded in .html files using <script> and <style> tags, or included via `<?!= include('filename') ?>` templating.

---

## 📊 SHEET STRUCTURE (existing)

### Member Directory (31 columns)
- Headers read dynamically — find columns by header name, never by index
- Key columns: Email, Role (steward/member/both), Name, Unit, Phone
- Role column values: "Steward", "Member", "Both"

### Grievance Log (28 columns)
- Key columns: Grievance ID, Member Email, Status, Current Step, Deadline, Assigned Steward
- Status values: New, Active, Overdue, Resolved
- Step values: Step 1, Step 2, Step 3

### Config Tab (existing — adding fields)
- Fields to add/ensure: Org Name, Org Abbreviation, Accent Hue (0-360), Magic Link Expiry Days, Cookie Duration Days, Logo Initial(s)

---

## 🔄 CHANGE LOG

### 2026-02-21 — Initial creation
- Created AI_REFERENCE.md
- Created project file structure
- Decisions locked: bound script, SSO + magic link auth, Mono Signal (steward) + Glass Depth (member), dynamic config, Bitly redirect, DDS-Dashboard only
- Auth: 7-day magic link, 30-day cookie, dual-role defaults to steward

---

## 🐛 ERRORS & FIXES

(none yet — will be logged here as they occur)

---

## ✅ FEATURES & HOW THEY WORK

### Auth System
- `doGet(e)` checks for: (1) valid cookie token, (2) SSO session, (3) magic link token in URL params
- If none → serve auth.html (login screen)
- Magic link: `Auth.gs` generates HMAC-signed token with email + timestamp, stores in ScriptProperties, emails via MailApp
- Cookie: set via URL param on redirect after successful auth, read via `e.parameter.token`
- Note: Apps Script web apps cannot set HTTP cookies directly. "Cookie" is implemented as a token stored in localStorage on the client side + validated against ScriptProperties on the server side.

### Role Routing
- After auth → email matched against Member Directory → Role column read
- "Steward" → steward.html, "Member" → member.html, "Both" → steward.html with toggle
- If email not found → error page with "Contact your steward" message

### Config Reader
- `ConfigReader.gs` reads Config tab, caches in CacheService (6 hour TTL)
- Returns: { orgName, orgAbbrev, accentHue, magicLinkExpiryDays, cookieDurationDays, logoInitials }
- Frontend receives config as JSON injected into page template

### Data Service
- `DataService.gs` exposes functions callable via `google.script.run`
- `getStewcardCases(stewardEmail)` — returns cases assigned to this steward
- `getMemberGrievances(memberEmail)` — returns only this member's grievances
- `getKPIs(stewardEmail)` — returns personal caseload stats
- All functions validate that the requesting user has permission to see the data

---

## 📝 NOTES FOR FUTURE LLMs

- If you're asked to modify this project, READ THIS FILE FIRST.
- Check the CHANGE LOG for recent modifications.
- All column lookups MUST be by header name, never by index.
- The Config tab is the single source of truth. Do not hardcode org-specific values.
- Test with at least two different org configs to verify dynamic behavior.
- The Bitly link means the deployment URL is not stable — never reference it in code or UI.
