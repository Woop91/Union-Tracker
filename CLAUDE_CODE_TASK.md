# Claude Code Task: Notifications Page Implementation

## Repo
- **Repo:** `Woop91/DDS-Dashboard` (private)
- **Branch:** `staging`
- **Token:** Read from AI_REFERENCE.md or use the one in git remote already configured

## READ FIRST
Before writing any code, read these files completely:
1. `AI_REFERENCE.md` — full system architecture, changelog, design decisions
2. `src/01_Core.gs` lines 737-800 (SHEETS constants) and lines 1796-1820 (NOTIFICATIONS_HEADER_MAP_)
3. `src/05_Integrations.gs` lines 1640-1810 (doGet routing) and lines 4060-4300 (existing notification API functions)
4. `src/05_Integrations.gs` — look at `getWebAppResourcesHtml()` (line ~3711) and `getWebAppCheckInHtml()` (line ~3895) for the EXACT pattern of how HTML page generators work in this codebase. Match their string concatenation style.

## CRITICAL ARCHITECTURE RULES
- **Everything is dynamic.** Column numbers come from HEADER_MAPs. Never hardcode column positions.
- **All HTML is returned as string concatenation** — GAS has no template engine. Use `var p = []; p.push(...); return p.join('');` pattern (same as getWebAppResourcesHtml uses string concat with `+`).
- **Do NOT use ES6** in .gs files — no arrow functions, no template literals, no let/const, no destructuring. Use `var`, `function(){}`, string concat with `+`.
- **Do NOT use heredocs or cat >> to write GAS code** — this causes escaping nightmares. Use proper file editing.
- Run `npm run build` after every change and verify it succeeds.
- Run `npm run test:unit` — there are 2 pre-existing test failures (architecture test + version mismatch). No new failures should be introduced.

## WHAT TO IMPLEMENT

### 1. Add `getNotificationRecipientListFull()` to `src/05_Integrations.gs`
Place it right after the existing `getNotificationRecipientList()` function (line ~4299).

**Purpose:** Returns full member list with directory columns for the steward compose form's recipient picker.

**Signature:** `function getNotificationRecipientListFull()`

**Returns:** `Object[]` — each object has: `{ name, email, location, department, jobTitle }`

**Implementation:**
- Read from `SHEETS.MEMBER_DIR`
- Use `MEMBER_COLS.FIRST_NAME`, `MEMBER_COLS.LAST_NAME`, `MEMBER_COLS.EMAIL` (required)
- Use `MEMBER_COLS.WORK_LOCATION`, `MEMBER_COLS.DEPARTMENT`, `MEMBER_COLS.JOB_TITLE` (optional — check if they exist before accessing)
- Sort by name alphabetically
- Wrap in try/catch, return `[]` on error, call `logError_()` on failure

### 2. Add `case 'notifications'` to doGet switch in `src/05_Integrations.gs`
Find the switch statement (~line 1746). Add between the `case 'resources'` and `case 'portal'` blocks:

```javascript
    case 'notifications':
      // v4.12.0: Notifications — members view/dismiss, stewards compose inline
      html = getWebAppNotificationsHtml();
      break;
```

### 3. Add `getWebAppNotificationsHtml()` to `src/05_Integrations.gs`
This is the main deliverable. Place it after `getNotificationRecipientListFull()`.

**This is a dual-role page:**
- **Members** see: notification list with dismiss buttons
- **Stewards** see: inline compose form at top + notification list below

**Role detection pattern** (copy from existing code):
```javascript
var isSteward = false;
var userEmail = '';
var userName = '';
try {
  var authResult = checkWebAppAuthorization('steward');
  isSteward = authResult.isAuthorized;
  userEmail = authResult.email || Session.getActiveUser().getEmail() || '';
  userName = userEmail.split('@')[0] || '';
} catch (authErr) {
  try { userEmail = Session.getActiveUser().getEmail() || ''; } catch (e2) { userEmail = ''; }
}
```

**Server-side data prefetch:**
```javascript
var userRole = isSteward ? 'steward' : 'member';
var notifications = getWebAppNotifications(userEmail, userRole);
var notifJson = JSON.stringify(notifications || []);
var recipientJson = isSteward ? JSON.stringify(getNotificationRecipientListFull() || []) : '[]';
```

**Visual design (MUST match existing pages):**
- Fonts: `DM Sans` (body) + `Fraunces` (serif display headers) — import from Google Fonts
- Hero gradient: `linear-gradient(145deg, #92400e, #b45309)` (amber/orange for notifications)
- Background: `#fafaf9` (warm off-white)
- Cards: white `#fff`, border `#f5f5f4`, shadow `0 1px 4px rgba(0,0,0,0.06)`, border-radius `14px`
- Compose form: white card, border-radius `16px`, shadow `0 2px 12px rgba(0,0,0,0.08)`, negative top margin to overlap hero
- See `getWebAppResourcesHtml()` and `getWebAppCheckInHtml()` for exact style patterns

**Steward compose form (inline at top of page) must have:**

A) **Recipient picker with two tabs:**
   - **Groups tab** (default): Three buttons — "All Members", "All Stewards", "Everyone". Click selects one (highlight with amber border).
   - **Individuals tab**: Shows filterable member list from `allMembers` (prefetched JSON).
     - **Filter bar** with:
       - Text search input (filters by name or email)
       - Location dropdown (populated from unique `location` values in member data)
       - Department dropdown (populated from unique `department` values)
       - Job Title dropdown (populated from unique `jobTitle` values)
     - **Scrollable member list** (max-height 240px):
       - Each row: checkbox + name + detail line (email · location · department)
       - Click row to toggle selection
       - Show selected count below list ("3 selected")
     - Filter dropdowns built client-side from the `allMembers` JSON on page load

B) **Form fields:**
   - Type: dropdown (Steward Message, Announcement, Deadline, System)
   - Priority: dropdown (Normal, Urgent)
   - Title: text input
   - Message: textarea
   - Expires: date input (blank = no auto-expiry)

C) **Send button:**
   - If Groups tab active: calls `sendWebAppNotification()` once with the group name as recipient
   - If Individuals tab active: calls `sendWebAppNotification()` once per selected member email
   - Show loading state on button, toast on success/error
   - After success: clear form, refresh notification list

**Notification list (both roles) must have:**
- Each notification renders as a card with:
  - Type badge (color-coded: blue for Steward Message, green for Announcement, red for Deadline, purple for System)
  - "URGENT" badge if priority is Urgent
  - Title (Fraunces serif, 15px, bold)
  - Message body
  - Meta line: "From: [sentBy]" and "Expires: [expiresDate]" if present
  - Created date in top-right
  - Dismiss button (✕) in top-right corner
- Dismiss calls `google.script.run.dismissWebAppNotification(id, USER_EMAIL)`
- On dismiss success: animate card out, remove from DOM
- Empty state: bell icon + "All caught up!" message

**Client-side JavaScript must:**
- Use `google.script.run` for all server calls (standard GAS web app pattern)
- Use `.withSuccessHandler()` and `.withFailureHandler()` 
- NOT use ES6 syntax (no arrow functions, no template literals, no const/let)
- Inject server data via: `var notifications = <?= notifJson ?>;` — WAIT, this codebase doesn't use templated HTML. Instead, inject as inline script variables in the HTML string:
  ```javascript
  p.push('<script>');
  p.push('var USER_EMAIL="' + userEmail.replace(/"/g, '\\"') + '";');
  p.push('var IS_STEWARD=' + String(isSteward) + ';');
  p.push('var notifications=' + notifJson + ';');
  p.push('var allMembers=' + recipientJson + ';');
  ```

### 4. Build, test, commit, push to staging

```bash
npm run build        # Must succeed. Current: ~62,500 lines / ~2,585 KB
npm run test:unit    # Only 2 pre-existing failures allowed (01_Core version + architecture)
git add -A
git commit -m "feat: add Notifications page with steward compose, recipient picker, dismiss (v4.12.0)

- New ?page=notifications route in doGet
- getWebAppNotificationsHtml() — dual-role page (member view + steward compose)
- getNotificationRecipientListFull() — member list with location/dept/title for filtering
- Steward inline compose: groups (All Members/Stewards/Everyone) + individual picker
- Individual picker: filterable by name, location, department, job title
- Member view: notification cards with dismiss, type badges, urgency indicators
- Notifications persist until steward-set expiry date OR individual member dismisses"
git push origin staging
```

### 5. Update AI_REFERENCE.md
Add to the v4.12.0 changelog section:
- Note the new route `?page=notifications`
- Note `getWebAppNotificationsHtml()` and `getNotificationRecipientListFull()` functions
- Add `?page=notifications` to the route table

### 6. Sync all branches
```bash
# Sync Main
git checkout Main
git merge staging -m "merge: sync staging v4.12.0 notifications to Main"
git push origin Main

# Sync dev
git checkout dev
git merge staging -m "merge: sync staging v4.12.0 notifications to dev"
git push origin dev

# Return to staging
git checkout staging
```

If merge conflicts occur on `dist/ConsolidatedDashboard.gs`:
```bash
git checkout --theirs dist/ConsolidatedDashboard.gs
npm run build
git add dist/ConsolidatedDashboard.gs
git merge --continue
# OR: git -c core.editor=true merge --continue
```

## VERIFICATION CHECKLIST
After all changes:
- [ ] `npm run build` succeeds
- [ ] `npm run test:unit` — only 2 pre-existing failures
- [ ] `grep "case 'notifications'" src/05_Integrations.gs` returns a match
- [ ] `grep "function getWebAppNotificationsHtml" src/05_Integrations.gs` returns a match
- [ ] `grep "function getNotificationRecipientListFull" src/05_Integrations.gs` returns a match
- [ ] `node -e "new Function(require('fs').readFileSync('src/05_Integrations.gs','utf-8'))"` — no syntax errors
- [ ] All 3 branches (staging, Main, dev) pushed and in sync
- [ ] AI_REFERENCE.md updated

## WHAT NOT TO DO
- Do NOT use `cat >>` or heredocs to append code — use proper file editing
- Do NOT use ES6 syntax anywhere in .gs files
- Do NOT delete or modify any existing routes, tabs, or functions
- Do NOT hardcode column numbers — always use MEMBER_COLS, NOTIFICATIONS_COLS, etc.
- Do NOT create separate CSS/JS files — everything is inline in the HTML string
- Do NOT use HtmlService.createTemplateFromFile — this codebase uses createHtmlOutput with string building
