# AI REFERENCE DOCUMENT — Grievance Dashboard Web App
# ⚠️ THIS FILE MUST NEVER BE DELETED. ONLY APPEND. ⚠️
# Used by: Claude, Gemini, ChatGPT, or any LLM working on this codebase.
# Last updated: 2026-02-21

---

## 🏗️ PROJECT OVERVIEW

- **What:** Mobile-optimized web app for workplace org grievance tracking & member management
- **Served by:** Google Apps Script doGet() → HtmlService, bound to a Google Sheet
- **Repo:** Woop91/MULTIPLE-SCRIPS-REPO → branch: web-dashboard
- **All other repos DEPRECATED** (509-dashboard, union-steward-app)
- **Access:** Users arrive via Bitly short link, never the raw script URL

---

## 🔴 CRITICAL RULES

1. **EVERYTHING DYNAMIC.** Never hardcode org names, column positions, sheet names.
2. **Config tab is HORIZONTAL** — headers in Row 1, values in Row 2+. Some columns are single-value, some are lists (dropdowns).
3. **Role detection:** "Is Steward" column = Yes/No. Stewards are always also members → role = "both". Non-stewards → role = "member".
4. **Bound to existing sheet** — same Apps Script project as the 15k-line dashboard.
5. **Bitly redirect** — deployment URL changes per deploy, Bitly stays constant.
6. **Auth: SSO + Magic Link** — both available to all users. 7-day magic link, 30-day session cookie.
7. **Dual-role users** default to steward view with toggle.
8. **Members see only their own grievances.** No PII of other members.
9. **Column lookups ALWAYS by header name**, never by index. Use aliases for flexibility.
10. **Assigned Steward** column may contain name OR email — code handles both.

---

## 📊 ACTUAL SHEET SCHEMA

### Member Directory (35 columns)
Member ID, First Name, Last Name, Job Title, Work Location, Unit, Cubicle,
Office Days, Email, Phone, Open Rate %, Preferred Communication,
Best Time to Contact, Supervisor, Manager, **Is Steward (Yes/No)**,
Committees, Assigned Steward, Last Virtual Mtg, Last In-Person Mtg,
Volunteer Hours, Interest: Local, Interest: Chapter, Interest: Allied,
Recent Contact Date, Contact Steward, Contact Notes,
Has Open Grievance?, Grievance Status, Days to Deadline,
Start Grievance, ⚡ Actions, PIN Hash, Employee ID, Department, Hire Date

### Grievance Log (40+ columns)
Grievance ID, Member ID, First Name, Last Name, Status, Current Step,
Incident Date, Filing Deadline, Date Filed,
**Step I Due, Step I Rcvd,**
**Step II Appeal Due, Step II Appeal Filed, Step II Due, Step II Rcvd,**
**Step III Appeal Due, Step III Appeal Filed,**
Date Closed, Days Open, Next Action Due, Days to Deadline,
Articles Violated, Issue Category, Member Email, Work Location,
Assigned Steward, Resolution, Message Alert, Coordinator Message,
Acknowledged By, Acknowledged Date, Drive Folder ID, Drive Folder URL,
⚡ Actions, Action Type, Checklist Progress,
Reminder 1 Date, Reminder 1 Note, Reminder 2 Date, Reminder 2 Note,
Last Updated

### Config Tab (HORIZONTAL — 60+ columns)
**Dropdown lists:** Job Titles, Office Locations, Units, Office Days, Yes/No,
Supervisors, Managers, Stewards, Steward Committees, Grievance Status,
Grievance Step, Issue Category, Articles Violated, Communication Methods,
Grievance Coordinators, Best Times to Contact, Home Towns, Unit Codes,
Escalation Statuses, Escalation Steps, Office Addresses,
Contract Article (Grievance/Discipline/Workload)

**Single values:** Organization Name, Local Number, Main Office Address,
Main Phone, Main Fax, Main Contact Name/Email, Chief Steward Email,
Google Drive Folder ID, Google Calendar ID, Grievance Form URL,
Contact Form URL, Admin Emails, Alert Days Before Deadline,
Notification Recipients, Filing Deadline Days, Step I Response Days,
Step II Appeal Days, Step II Response Days, Contract Name, Union Parent,
State/Region, Organization Website, Satisfaction Survey URL,
Archive Folder ID, Template ID, PDF Folder ID, 📱 Mobile Dashboard URL

**Dashboard-specific (to add):** Accent Hue, Logo Initials,
Magic Link Expiry Days, Cookie Duration Days, Steward Label, Member Label

---

## 🎨 DESIGN SYSTEM

### Steward — "Mono Signal"
- JetBrains Mono body, Space Grotesk display
- 10px radius, dense layout, dot status indicators
- Dark default, light toggle available

### Member — "Glass Depth"
- Plus Jakarta Sans body, Sora display
- 20px radius, frosted glass cards, spacious layout
- Dark default, light toggle available

### Accent Color
- Stored as hue (0-360) in Config tab "Accent Hue" column
- Full palette generated dynamically via HSL

---

## 📁 FILE STRUCTURE

WebApp.gs, Auth.gs, ConfigReader.gs, DataService.gs,
index.html, styles.html, auth_view.html, steward_view.html,
member_view.html, error_view.html

---

## 🔄 CHANGE LOG

### 2026-02-21 — Initial build
- Created full Phase 1: auth, config, data service, both views
- Schema aligned to actual Member Directory (35 cols), Grievance Log (40+ cols), Config (60+ cols)
- Role detection via "Is Steward" = Yes/No (not a separate Role column)
- Config tab reads horizontally (headers in Row 1, values in Row 2+)
- Timeline built from actual step dates (Step I Due/Rcvd, Step II Appeal Due/Filed, etc.)
- Assigned Steward matching handles both email and name lookups
- PIN Hash column exists but not yet used for auth (magic link + SSO used instead)

---

## 🐛 ERRORS & FIXES

(none yet)

---

## 📌 CONFIRMED DECISIONS (2026-02-21)

- "Assigned Steward" in Grievance Log stores **names, not emails**. Steward lookup does name match against Member Directory.
- **PIN Hash** column will be used as a third auth option (Phase 2): SSO, magic link, or PIN.
- "File a Grievance" button links to `config.grievanceFormUrl` from Config tab.
- "📱 Mobile Dashboard URL" in Config tab stores the **Bitly short link** for easy sharing. Not referenced by the app code itself.
- Dashboard-specific columns (Accent Hue, Logo Initials, etc.) already added to Config tab by user.

---

## ✅ FEATURES

### Auth System
- Google SSO: Session.getActiveUser().getEmail()
- Magic Link: 7-day token via MailApp, stored in ScriptProperties
- Session: 30-day token in localStorage, validated server-side
- Token cleanup: cleanupExpiredTokens() for time-based trigger

### Role Detection
- "Is Steward" Yes → role "both" (steward view with member toggle)
- "Is Steward" No → role "member"
- Email not found → error page with "contact your steward"

### Timeline (Member View)
- Built from actual grievance step dates, not generated
- Shows: Filed → Step I Due/Rcvd → Step II Appeal/Response → Step III → Closed
- Done/pending states from whether date columns have values

### Steward Matching
- "Assigned Steward" column may contain email or full name
- DataService tries email lookup first, falls back to name search

---

## 📝 NOTES FOR FUTURE LLMs

- The Config tab is HORIZONTAL. Row 1 = headers, Row 2+ = values/lists.
- "Is Steward" is a Yes/No field, NOT a role enum.
- PIN Hash column exists — future auth enhancement possibility.
- Step dates are separate columns (Step I Due, Step I Rcvd, etc.) — use these for timeline.
- "Days to Deadline" exists as a calculated column in both sheets — use it if available, else calculate.
- The "Reminder 2 Note" column has a typo in the sheet ("eminder 2 Note") — code handles this alias.
