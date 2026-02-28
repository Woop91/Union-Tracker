# Strategic Command Center - Complete Features Reference

**Version:** 4.13.0 | **Codename:** SPA Web Dashboard & Notifications
**Last Updated:** February 2026

> **New in v4.13.0:** Notification bell with unread badge, EventBus auto-notifications, WorkloadService SPA module, individual-file build system

This document provides a comprehensive, searchable reference of all features in the Dashboard system. Use `Ctrl+F` (or `Cmd+F` on Mac) to search for specific features.

---

## Table of Contents

1. [Dashboard & Analytics](#1-dashboard--analytics)
2. [Search & Discovery](#2-search--discovery)
3. [Grievance Management](#3-grievance-management)
4. [Member Management](#4-member-management)
5. [Steward Tools](#5-steward-tools)
6. [Calendar & Scheduling](#6-calendar--scheduling)
7. [Google Drive Integration](#7-google-drive-integration)
8. [Notifications & Alerts](#8-notifications--alerts)
9. [Accessibility & Comfort View](#9-accessibility--comfort-view)
10. [Strategic Intelligence](#10-strategic-intelligence)
11. [Data Validation & Integrity](#11-data-validation--integrity)
12. [Administration & Maintenance](#12-administration--maintenance)
13. [Mobile & Field Access](#13-mobile--field-access)
14. [Web App & Portal](#14-web-app--portal)
15. [Security & Audit](#15-security--audit)
16. [Document Generation](#16-document-generation)
17. [Demo & Development Tools](#17-demo--development-tools)
18. [Dynamic Engine & Extensions](#18-dynamic-engine--extensions)
19. [Grievance Reminders](#19-grievance-reminders)
20. [Looker Studio Integration](#20-looker-studio-integration)
21. [Member Self-Service & PIN Authentication](#21-member-self-service--pin-authentication)
22. [Meeting Check-In & Document Automation](#22-meeting-check-in--document-automation)
23. [Member Drive Folders](#23-member-drive-folders)
24. [Constant Contact Integration](#24-constant-contact-integration)
25. [Workload Tracker](#25-workload-tracker)
26. [Resources Hub](#26-resources-hub)
27. [Notifications System](#27-notifications-system)
28. [SPA Web Dashboard](#28-spa-web-dashboard)
29. [Notification Bell & EventBus Alerts](#29-notification-bell--eventbus-alerts)

---

## 1. Dashboard & Analytics

### Two-Dashboard Architecture (v4.3.3)

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Steward Dashboard** | Internal dashboard with 11 tabs: Overview, My Cases, Workload, Analytics, Directory, Hot Spots, Bargaining, Satisfaction, Resources, Compare, Meeting Notes. Contains member names and PII - for steward use only. | Strategic Ops > Command Center > Steward Dashboard | internal, analytics, PII, workload |
| **Member Dashboard** | PII-safe dashboard for sharing with members. Shows aggregate stats, steward directory, and satisfaction scores without personal information. | Strategic Ops > Command Center > Member Dashboard | public, aggregate, safe, sharing |
| **Executive Command Dashboard** | Legacy 5-tab modal with Overview, My Cases, Grievances, Members, and Analytics tabs. | Union Hub > Dashboards > Dashboard | executive, overview, legacy |
| **Interactive Dashboard** | Customizable dashboard with 20+ metrics and 7 chart types. | See INTERACTIVE_DASHBOARD_GUIDE.md | charts, metrics, customizable |

### Dashboard Tabs (Steward Dashboard)

| Tab | Description | Key Metrics |
|-----|-------------|-------------|
| **Overview** | High-level organization health metrics with Quick Insights panel | Active grievances, win rate, member count, morale score, engagement summary, bargaining position |
| **My Cases** | Steward's assigned grievances (PII mode only) | Active cases, urgent count, avg days open, filtering by status |
| **Workload** | Steward case distribution and capacity | Cases per steward, overload warnings (8+ cases), top performers |
| **Analytics** | Grievance trends and patterns | Status breakdown, engagement metrics, volunteer hours |
| **Directory** | Member contact trends and data quality | Recent updates, stale contacts, missing email/phone |
| **Hot Spots** | Problem areas with 4 hot spot types | Grievance clusters, dissatisfaction areas, low engagement zones, overdue concentrations |
| **Bargaining** | Comprehensive bargaining intelligence | Step 1/2 denial rates, step outcomes, case details at each step, recent grievances |
| **Satisfaction** | 8-section member satisfaction with individual question scores | Section scores, question breakdowns, expandable details |
| **Resources** | Organization documents and steward contacts | Google Drive folder, steward directory with search |
| **Compare** | Dashboard comparison and export tool | Period comparison, step-by-step breakdown, satisfaction comparison, denial rate analysis, CSV export |
| **Meeting Notes** | Chronological meeting notes with view-only links (v4.6.0) | Meeting name, date, search, Google Doc view links |
| **Help** | FAQ and documentation (PII mode only) | Searchable help with 4 categories and 12+ questions |

---

## 2. Search & Discovery

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Desktop Search** | Advanced search with tabs for All/Members/Grievances, filters by status/department/date, and result previews. | Strategic Ops > Desktop Search | advanced, filter, desktop |
| **Quick Search** | Minimal interface for fast member/grievance lookup. Supports partial name matching. | Union Hub > Quick Search | fast, simple, minimal |
| **Advanced Search** | Fullscreen search with complex filtering and multiple criteria support. | Union Hub > Search > Advanced Search | fullscreen, complex, multi-criteria |
| **Mobile Search** | Touch-optimized search for field use on phones and tablets. | Field Portal > Mobile Search | mobile, touch, field |
| **Live Steward Search** | Real-time client-side filtering of steward directory in Member Dashboard. | Within Member Dashboard | real-time, instant, filter |
| **Searchable Help Guide** | 4-tab help modal with real-time search across Overview, Menu Reference, FAQ, and Quick Tips. | Union Hub > Help & Documentation | help, FAQ, documentation |

---

## 3. Grievance Management

### Core Functions

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **New Case/Grievance** | Opens pre-filled grievance form. Calculates deadlines automatically based on Article 23A. | Strategic Ops > Cases > New Case/Grievance | create, file, new |
| **Edit Selected Grievance** | Modify details of the currently selected grievance row. | Strategic Ops > Cases > Edit Selected | edit, modify, update |
| **Advance Grievance Step** | Move grievance to next step (Step I → Step II → Step III → Arbitration). | Strategic Ops > Cases > Advance Step | step, advance, escalate |
| **View Active Grievances** | Filter to show only open/pending grievances. | Union Hub > Grievances > View Active | active, open, pending |
| **Bulk Status Update** | Update status for multiple selected grievances at once. | Union Hub > Grievances > Bulk Update | bulk, batch, multiple |

### Deadline Management

| Feature | Description | Keywords |
|---------|-------------|----------|
| **Auto-Calculated Deadlines** | Based on Article 23A: Step 1 (7 days), Step 2 Appeal (7 days), Step 2 Response (14 days), Step 3 Appeal (10 days), Step 3 Response (21 days), Arbitration (30 days). | deadline, calculate, Article 23A |
| **Days to Deadline** | Auto-updating column showing days remaining or "Overdue" status. | countdown, overdue, remaining |
| **Message Alert Flag** | Checkbox to highlight urgent cases in yellow and move to top when sorted. | urgent, flag, priority |
| **Next Action Due** | Shows which deadline is coming next for each grievance. | next, action, due |

### Grievance Analytics

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Grievance Trends** | Analyze filing patterns over time. | Field Portal > Analytics > Grievance Trends | trends, patterns, history |
| **Search Precedents** | Find similar past grievances for reference. | Field Portal > Analytics > Search Precedents | precedent, similar, past |
| **Status Breakdown** | Pie chart of grievance statuses (Open, Won, Denied, etc.). | Within Dashboards | status, pie, breakdown |
| **Step Distribution** | Analysis of where grievances are in the process. | Within Dashboards | step, distribution, pipeline |

---

## 4. Member Management

### Core Functions

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Add New Member** | Open member registration form with all fields. | Union Hub > Members > Add New Member | add, register, new |
| **Find Member** | Search for specific member by name, ID, or other criteria. | Union Hub > Members > Find Member | find, search, lookup |
| **Import Members** | Bulk import member data from external sources. | Union Hub > Members > Import Members | import, bulk, external |
| **Export Members** | Export member directory to CSV or other formats. | Union Hub > Members > Export Members | export, CSV, download |
| **Steward Directory** | View list of all stewards with contact information. | Union Hub > Members > Steward Directory | steward, directory, contact |

### Member ID System

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Auto-Generate Member IDs** | Creates name-based IDs from member names (e.g., MJASM472 for Jane Smith). | Strategic Ops > ID Engines > Generate Missing IDs | ID, generate, auto |
| **Check Duplicate IDs** | Finds and highlights duplicate Member IDs in the directory. | Strategic Ops > ID Engines > Check Duplicates | duplicate, check, validate |
| **Unit Code Configuration** | Custom unit codes for ID generation (e.g., Main Station:MS). | Config sheet column AT | unit, code, config |

### Member Directory Columns

| Column | Description | Auto-Calculated? |
|--------|-------------|------------------|
| **Member ID** | Unique identifier (name-based format, e.g., MJASM472) | Optional auto-gen |
| **Has Open Grievance** | Yes/No based on Grievance Log | Yes |
| **Grievance Status** | Current status if has open grievance | Yes |
| **Days to Deadline** | Days remaining on current grievance | Yes |
| **Assigned Steward** | Steward responsible for this member | No |
| **Start Grievance** | Checkbox to open pre-filled grievance form | Resets after use |

---

## 5. Steward Tools

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Promote to Steward** | Change member status to steward, sends toolkit email. | Strategic Ops > Steward Management > Promote | promote, steward, new |
| **Demote from Steward** | Remove steward status from member. | Strategic Ops > Steward Management > Demote | demote, remove, former |
| **Steward Performance Modal** | View active cases, total cases, and win rates for all stewards. | Strategic Ops > Command Center > Steward Performance | performance, win rate, cases |
| **Steward Contact Forms** | Send contact form links to stewards. | Union Hub > Members > Steward Contact Forms | contact, form, steward |
| **Steward Workload Report** | Capacity analysis with overload detection (flags 8+ active cases). | Strategic Ops > Analytics > Workload Report | workload, capacity, overload |
| **Rising Stars** | Highlights top-performing stewards by score and win rate. | Strategic Ops > Strategic Intelligence > Rising Stars | top, performance, best |

---

## 6. Calendar & Scheduling

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Sync Deadlines to Calendar** | Creates Google Calendar events for all grievance deadlines. | Union Hub > Calendar > Sync Deadlines | sync, calendar, events |
| **View Upcoming Deadlines** | Shows calendar of upcoming deadlines. | Union Hub > Calendar > View Upcoming | upcoming, calendar, view |
| **Clear Calendar Events** | Remove previously synced calendar events. | Union Hub > Calendar > Clear Events | clear, remove, delete |
| **Create Calendar Event** | Manually create calendar event for selected grievance. | Dashboard > Calendar > Create Event | create, event, manual |

---

## 7. Google Drive Integration

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Setup Folder for Grievance** | Auto-creates Drive folder with subfolders for each step. | Union Hub > Google Drive > Setup Folder | folder, drive, create |
| **View Grievance Files** | Open the Drive folder for selected grievance. | Union Hub > Google Drive > View Files | view, files, open |
| **Batch Create Folders** | Create Drive folders for multiple grievances at once. | Union Hub > Google Drive > Batch Create | batch, bulk, folders |

### Folder Structure Created

```
[Grievance ID] - [Member Name]/
├── Step I/
├── Step II/
├── Step III/
├── Arbitration/
└── Supporting Documents/
```

---

## 8. Notifications & Alerts

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Email Notifications** | Send email alerts for deadlines and status changes. | Union Hub > Notifications | email, alert, notify |
| **Deadline Reminders** | Automatic reminders before deadlines. | Triggered automatically | reminder, deadline, auto |
| **Escalation Alerts** | Notifications when cases reach Step II/III/Arbitration. | Config column AV-AW | escalation, step, alert |
| **Overdue Alerts** | Daily alerts for overdue grievances (via midnight trigger). | Automatic at midnight | overdue, alert, daily |
| **Test Notification** | Send test notification to verify setup. | Union Hub > Notifications > Test | test, verify, setup |

---

## 9. Accessibility & Comfort View

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Focus Mode** | Distraction-free view, hides non-essential sheets. | Union Hub > Comfort View > Focus Mode | focus, distraction-free, ADHD |
| **Zebra Stripes** | Alternating row colors for easier reading. | Union Hub > Comfort View > Zebra Stripes | zebra, stripes, alternating |
| **High Contrast Mode** | Enhanced contrast for visibility. | Union Hub > Comfort View > High Contrast | contrast, visibility, accessibility |
| **Reduced Motion** | Minimizes animations for motion sensitivity. | Union Hub > Comfort View > Reduced Motion | motion, animation, sensitivity |
| **Dark Mode** | Dark gradient backgrounds across all modals. | Union Hub > View > Dark Mode | dark, theme, night |
| **Apply Global Styling** | Applies Roboto font and zebra stripes to all rows in data sheets. | Dashboard > Styling > Apply Global | styling, font, Roboto |
| **ADHD-Friendly Themes** | Color schemes designed for attention difficulties. | See COMFORT_VIEW_GUIDE.md | ADHD, theme, friendly |

---

## 10. Strategic Intelligence

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Unit Hot Zones** | Identifies locations with 3+ active grievances. | Strategic Ops > Strategic Intelligence > Hot Zones | hot zones, problem areas, locations |
| **Rising Stars** | Top-performing stewards by score and win rate. | Strategic Ops > Strategic Intelligence > Rising Stars | top performers, stewards, win rate |
| **Management Hostility Report** | Analyzes denial rates across grievance steps. | Strategic Ops > Strategic Intelligence > Hostility Report | denial, management, hostility |
| **Bargaining Cheat Sheet** | Strategic data for contract negotiations (common violations, leverage points). | Strategic Ops > Strategic Intelligence > Bargaining | bargaining, contract, negotiation |
| **Unit Density Treemap** | Visual heat map of grievance activity by unit (green → yellow → red). | Strategic Ops > Analytics > Treemap | treemap, density, heatmap |
| **Sentiment Trend Analysis** | Organization morale tracking over time from survey data. | Strategic Ops > Analytics > Sentiment Trends | sentiment, morale, trends |
| **Unit Health Report** | Comprehensive health check for specific units. | Field Portal > Analytics > Unit Health | health, unit, report |

---

## 11. Data Validation & Integrity

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Run Bulk Validation** | Validate all data for consistency and errors. | Admin > Validation > Run Bulk Validation | validate, bulk, check |
| **Validation Settings** | Configure validation rules and thresholds. | Admin > Validation > Settings | settings, rules, configure |
| **Validation Indicators** | Visual indicators for validation status. | Admin > Validation > Indicators | indicators, status, visual |
| **Orphan Record Detection** | Find grievances without matching members. | Within System Diagnostics | orphan, mismatch, integrity |
| **Setup Data Validations** | Apply dropdown validations to all cells. | Admin > Setup > Data Validations | dropdown, validation, setup |

---

## 12. Administration & Maintenance

### Diagnostics & Repair

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **System Diagnostics** | Comprehensive health check on all components. | Admin > System Diagnostics | diagnostics, health, check |
| **Repair Dashboard** | Auto-fix common issues (missing sheets, broken formulas). | Admin > Repair Dashboard | repair, fix, auto |
| **Refresh All Formulas** | Recalculate all formulas in the spreadsheet. | Settings > Refresh All Formulas | refresh, formulas, recalculate |

### Automation & Triggers

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Install Auto-Sync Trigger** | Real-time data synchronization on edits. | Admin > Data Sync > Install Trigger | sync, trigger, auto |
| **Midnight Auto-Refresh** | Daily 12AM trigger refreshes dashboards and sends overdue alerts. | Admin > Automation > Midnight Trigger | midnight, daily, refresh |
| **Email Snapshots** | Scheduled email reports with dashboard data. | Admin > Automation > Email Snapshots | email, snapshot, scheduled |
| **Remove All Triggers** | Clean up all automation triggers. | Admin > Triggers > Remove All | remove, cleanup, triggers |

### Cache & Performance

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Cache Status** | View current cache utilization. | Admin > Cache > Status | cache, status, view |
| **Warm Up Caches** | Pre-load frequently accessed data. | Admin > Cache > Warm Up | warm, preload, performance |
| **Clear All Caches** | Reset all cached data. | Admin > Cache > Clear All | clear, reset, cache |

### Setup Functions

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Setup All Hidden Sheets** | Initialize all 6 calculation sheets. | Admin > Setup > Hidden Sheets | hidden, setup, initialize |
| **Setup Data Validations** | Apply all dropdown validations. | Admin > Setup > Data Validations | validation, dropdown, setup |
| **Apply Default Settings** | Reset to default configuration. | Admin > Setup > Default Settings | default, reset, settings |

---

## 13. Mobile & Field Access

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Pocket/Mobile View** | Hides non-essential columns for phone access. | Dashboard > Field Access > Mobile View | mobile, pocket, phone |
| **Get Mobile URL** | Copy the spreadsheet URL for mobile access. | Field Portal > Field Accessibility > Get URL | URL, mobile, share |
| **Mobile-Optimized Search** | Touch-friendly search interface. | Field Portal > Mobile Search | touch, mobile, search |
| **Quick Actions** | Simplified actions for common field tasks. | Within Mobile View | quick, simple, field |

---

## 14. Web App & Portal

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Deploy Web App** | Create standalone web application from the dashboard. | Field Portal > Web App > Deploy | deploy, web app, standalone |
| **Member Portal** | Personalized member view via URL parameter (?member=ID). | Via Web App URL | portal, personal, member |
| **Public Statistics Portal** | Aggregate statistics for public deployment. | Field Portal > Web App > Public Portal | public, statistics, aggregate |
| **Email Portal Links** | Send personalized dashboard URLs to members. | Field Portal > Web App > Email Links | email, portal, personalized |
| **JSON API Endpoints** | REST API for external integrations. | Via doGet/doPost handlers | API, JSON, REST |

---

## 15. Security & Audit

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Audit Logging** | Track all changes with timestamps and user info. | Automatic (hidden _Audit_Log sheet) | audit, log, tracking |
| **Sabotage Protection** | Detects mass deletion (>15 cells) and alerts. | Automatic on edit | sabotage, protection, deletion |
| **Safety Valve PII Scrubbing** | Auto-redacts phone numbers and SSN patterns from public dashboards. | Automatic in Member Dashboard | PII, scrub, redact, privacy |
| **Weingarten Rights Utility** | Emergency rights statement with tap-to-expand for member protection. | Within Member Dashboard | Weingarten, rights, legal |
| **Version History** | Access Google Sheets' built-in version history. | File > Version history | undo, history, restore |

---

## 16. Document Generation

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Create Grievance PDF** | Generate signature-ready PDF with legal blocks. | Strategic Ops > ID Engines > Create PDF | PDF, generate, signature |
| **Email PDF Attachment** | Send generated PDFs via email. | After PDF generation | email, PDF, send |
| **Batch PDF Generation** | Create PDFs for multiple grievances. | Strategic Ops > ID Engines > Batch PDF | batch, PDF, bulk |

---

## 17. Demo & Development Tools

> **Note:** These features are in `07_DevTools.gs` and should be removed before production deployment.

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Seed All Sample Data** | Generate 1,000 test members and 300 grievances. | Admin > Demo Data > Seed All | seed, demo, test |
| **Seed Members Only** | Generate only test member data. | Admin > Demo Data > Seed Members | seed, members, test |
| **Seed Grievances Only** | Generate only test grievance data. | Admin > Demo Data > Seed Grievances | seed, grievances, test |
| **NUKE Seeded Data** | Remove all demo data (3-5 minute warnings, preserves real data). | Admin > Demo Data > NUKE | nuke, delete, cleanup |
| **Developer Panel** | Access developer tools and debugging utilities. | Admin > Demo Data > Developer Panel | developer, debug, tools |

---

## Quick Reference: Menu Structure

### Union Hub (Main Menu)
- Search (Desktop, Quick, Advanced)
- Grievances (New, Edit, Bulk Update, View Active)
- Members (Add, Find, Import, Export, Steward Directory)
- Calendar (Sync, View, Clear)
- Google Drive (Setup, View, Batch)
- Notifications (Settings, Test)
- View (Dashboards, Dark Mode)
- Comfort View (Focus, Zebra, Contrast)
- Multi-Select Panel
- Help & Documentation

### Admin Menu
- System Diagnostics
- Repair Dashboard
- Automation (Triggers, Snapshots)
- Data Sync (Sync All, Install Triggers)
- Validation (Run, Settings, Indicators)
- Cache (Status, Warm Up, Clear)
- Setup (Hidden Sheets, Validations, Defaults)
- Demo Data (Seed, NUKE)

### Strategic Ops Menu
- Command Center (Steward Dashboard, Member Dashboard, Performance)
- Cases (New, Edit, Advance)
- Strategic Intelligence (Hot Zones, Rising Stars, Hostility, Bargaining)
- Analytics (Treemap, Sentiment, Workload)
- ID Engines (Generate IDs, Check Duplicates, PDF)
- Steward Management (Promote, Demote, Contact)

### Field Portal Menu
- Field Accessibility (Mobile View, Get URL)
- Analytics (Unit Health, Trends, Precedents)
- Web App (Deploy, Portals, Email Links)

---

## 18. Dynamic Engine & Extensions

> **File:** `12_Features.gs`

The Dynamic Engine provides extensible features for organizational structure tracking, custom columns, and self-healing formulas.

### Member Leaders (Organizational Layer)

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Get Member Leaders** | Retrieves all members marked as stewards or member leaders with role/unit info | `getMemberLeaders()` | leaders, stewards, organizational |
| **Stewards for Grievance** | Returns dropdown-ready list of stewards for grievance assignment | `getStewardsForGrievance()` | dropdown, assignment, steward |
| **Check Member Leader Status** | Determines if a specific member is a steward or member leader | `isMemberLeader(memberId)` | check, status, role |

### Column Expansion (No-Code Custom Columns)

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Get Header Map** | Retrieves column headers with 5-minute caching for performance | `getHeaderMap(sheetName)` | headers, cache, columns |
| **Expansion Column Data** | Gets data from custom/expansion columns by header name | `getExpansionColumnData(sheetName, columnHeader)` | custom, data, expansion |
| **Generate Expansion HTML** | Creates HTML form fields for custom column editing | `generateExpansionFieldsHtml(sheetName, rowIndex)` | form, HTML, custom fields |
| **Save Expansion Data** | Batch-saves data to expansion columns with smart contiguous detection | `saveExpansionData(sheetName, rowIndex, columnData)` | save, batch, custom |
| **Invalidate Header Cache** | Clears cached header maps when columns change | `invalidateHeaderCache(sheetName)` | cache, clear, refresh |

### Self-Healing Hidden Architecture

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Repair Dynamic Formulas** | Injects/repairs formulas in hidden `_Dashboard_Calc` sheet (row 50+) | `repairDynamicFormulas()` | repair, formulas, self-healing |
| **Setup Dynamic Engine** | Initializes all Dynamic Engine features and returns status | `setupDynamicEngine()` | setup, initialize, status |
| **Get Dynamic Engine Status** | Returns current status of all engine components | `getDynamicEngineStatus()` | status, health, diagnostics |

### Performance Optimizations

| Feature | Description |
|---------|-------------|
| **CacheService Layer** | 5-minute TTL caching for header maps |
| **Unified Data Loader** | Single-pass data reading via `loadMemberData_(options)` |
| **Pre-computed Column Indices** | `COL_IDX` object for fast array access |
| **Batch Writes** | Uses `setValues()` for efficient multi-cell updates |
| **Smart Contiguous Detection** | Optimizes saves for adjacent columns |

---

## 19. Grievance Reminders

> **File:** `12_Features.gs` (Reminders section)

Allows users to set two reminder dates with notes for scheduling meetings and follow-ups on grievances.

### Reminder Management

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Set Reminder** | Sets a reminder (1 or 2) with date and note for a grievance | `setGrievanceReminder(grievanceId, reminderNum, date, note)` | set, schedule, meeting |
| **Get Reminders** | Retrieves both reminders for a specific grievance | `getGrievanceReminders(grievanceId)` | get, view, reminders |
| **Clear Reminder** | Clears a specific reminder for a grievance | `clearGrievanceReminder(grievanceId, reminderNum)` | clear, remove, cancel |
| **Get Due Reminders** | Finds all grievances with reminders due within specified days | `getDueReminders(daysAhead)` | due, upcoming, alerts |
| **Reminder Summary** | Dashboard-ready summary with today/week counts | `getReminderSummary()` | summary, dashboard, counts |

### Reminder Dialog

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Show Reminder Dialog** | Opens dark-themed modal to manage reminders for selected grievance | `showReminderDialog(grievanceId)` | dialog, modal, UI |

### Automated Notifications

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Check & Notify** | Checks for due reminders and shows toast notifications | `checkAndNotifyReminders(daysAhead)` | notify, alert, check |
| **Install Trigger** | Sets up daily 8 AM trigger for automatic reminder checks | `installReminderTrigger()` | trigger, automation, daily |

### Grievance Log Columns (AL-AO)

| Column | Letter | Description |
|--------|--------|-------------|
| **Reminder 1 Date** | AL | First reminder date |
| **Reminder 1 Note** | AM | First reminder note (e.g., "Schedule Step II meeting") |
| **Reminder 2 Date** | AN | Second reminder date |
| **Reminder 2 Note** | AO | Second reminder note |

---

## 20. Looker Studio Integration

> **File:** `12_Features.gs` (Looker section)

Provides read-only data layers for Google Looker Studio without modifying existing sheets or modal dashboards.

### Standard Integration (With PII)

Creates hidden sheets for internal Looker reports with full member data.

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Setup Integration** | Creates hidden `_Looker_*` sheets for Looker data sources | `setupLookerIntegration()` | setup, create, initialize |
| **Refresh Data** | Updates all Looker data sheets from source data | `refreshLookerData()` | refresh, update, sync |
| **Connection Help** | Shows dialog with Looker Studio connection instructions | `showLookerConnectionHelp()` | help, connect, instructions |
| **Get Status** | Returns status of all Looker sheets and record counts | `getLookerStatus()` | status, health, counts |
| **Install Trigger** | Sets up daily 6 AM auto-refresh trigger | `installLookerRefreshTrigger()` | trigger, automation, daily |

#### Standard Looker Sheets (Restricted Sources)

| Sheet Name | Source | Data Included |
|------------|--------|---------------|
| `_Looker_Members` | Member Directory | Full member data with names, contact info, grievance stats |
| `_Looker_Grievances` | Grievance Log | Full grievance data with member names, dates, outcomes |
| `_Looker_Satisfaction` | Member Satisfaction | Anonymous survey responses with section averages (no email, no member ID — verification via vault hashes only) |

### PII-Free Integration (Anonymized)

Creates anonymized sheets for external stakeholders, public dashboards, or compliance reporting.

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Setup Anon Integration** | Creates hidden `_Looker_Anon_*` sheets with no PII | `setupLookerAnonIntegration()` | setup, anonymous, PII-free |
| **Refresh Anon Data** | Updates all anonymized Looker sheets | `refreshLookerAnonData()` | refresh, anonymous, sync |
| **Anon Connection Help** | Shows connection help for PII-free sheets | `showLookerAnonConnectionHelp()` | help, anonymous, connect |
| **Get Anon Status** | Returns status of anonymized Looker sheets | `getLookerAnonStatus()` | status, anonymous, health |
| **Install All Triggers** | Sets up combined refresh for both standard and PII-free | `installLookerAllRefreshTrigger()` | trigger, both, combined |

#### PII-Free Looker Sheets

| Sheet Name | Source | Data Excluded |
|------------|--------|---------------|
| `_Looker_Anon_Members` | Member Directory | Names, emails, phones, member IDs (uses hashes) |
| `_Looker_Anon_Grievances` | Grievance Log | Member names, member IDs, steward names |
| `_Looker_Anon_Satisfaction` | Member Satisfaction | Any potential PII linkages |

### Anonymization Features

| Feature | Description |
|---------|-------------|
| **Anonymous Hashes** | Non-reversible hash generation for IDs (e.g., `A7F3K2M9`) |
| **Bucketed Values** | Days since contact → "Within Week", "1-3 Months", etc. |
| **Categorized Roles** | Specific job titles → "Nursing", "Technical/Support", etc. |
| **Engagement Levels** | Calculated from aggregated factors, not individual data |
| **Score Buckets** | 1-10 ratings → "Low (1-4)", "Medium (5-7)", "High (8-10)" |

### Use Cases

| Integration Type | Best For |
|-----------------|----------|
| **Standard** | Internal steward dashboards, executive reports, detailed analytics |
| **PII-Free** | External stakeholders, public dashboards, compliance reporting, board presentations |

---

## 21. Member Self-Service & PIN Authentication

> **File:** `13_MemberSelfService.gs`

Provides PIN-based member authentication for self-service access to the dashboard system.

### PIN Management

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Generate Member PIN** | Creates secure UUID-based PIN using Utilities.getUuid() with hashed storage | `generateMemberPIN(memberId)` | PIN, generate, secure, UUID |
| **Authenticate Member** | Validates member-provided PIN against stored hash | `authenticateMember(memberId, pin)` | authenticate, login, verify |
| **Reset Member PIN** | Generates a new PIN and invalidates the previous one | `resetMemberPIN(memberId)` | reset, regenerate, new PIN |

### Self-Service Portal

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Member Self-Service View** | PIN-authenticated portal for members to view their own data | `showMemberSelfService()` | portal, self-service, view |
| **Member Contact Update** | Members can update Email, Phone, Preferred Contact, Best Time, and State from the portal | `updateMemberContact()` | contact, update, edit, state |
| **PIN Entry Dialog** | Secure PIN entry interface for member authentication | `showPINEntryDialog()` | dialog, PIN, entry |

---

## 22. Meeting Check-In & Document Automation

> **File:** `14_MeetingCheckIn.gs`, `05_Integrations.gs`

Comprehensive meeting management with document creation, scheduled notifications, and a member-facing meeting notes dashboard.

### Meeting Setup & Check-In

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Setup Meeting Dialog** | Creates meetings with name, date, time, type, duration, steward notification settings, and agenda steward selection | `showSetupMeetingDialog()` | setup, create, meeting |
| **Create Meeting** | Creates meeting row, Google Calendar event, and auto-generates Meeting Notes and Agenda Google Docs | `createMeeting()` | create, calendar, docs |
| **Check-In Members** | Email or PIN-based member check-in for meetings | `showMeetingCheckInDialog()` | check-in, attendance, email |
| **Meeting Event Scheduling** | Full calendar lifecycle: creates events, activates/deactivates check-in based on event status | Via `dailyTrigger()` | calendar, scheduling, lifecycle |

### Meeting Document Automation (v4.6.0)

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Auto-Create Meeting Notes Doc** | Google Doc created in `Meeting Notes/` folder when meeting is set up | `createMeetingDocs()` | notes, document, auto-create |
| **Auto-Create Meeting Agenda Doc** | Google Doc created in `Meeting Agenda/` folder when meeting is set up | `createMeetingDocs()` | agenda, document, auto-create |
| **Two-Tier Agenda Sharing** | Selected stewards get agenda 3 days before; ALL stewards get it 1 day before. Agenda never shared with members | `processMeetingDocNotifications()` | agenda, steward, sharing |
| **Meeting Notes Notification** | Stewards in NOTIFY_STEWARDS receive the notes link 1 day before meeting | `processMeetingDocNotifications()` | notes, notification, email |
| **View-Only Publishing** | Meeting Notes auto-set to view-only 1 day after the meeting | `setDocViewOnlyByLink()` | publish, view-only, sharing |
| **Steward Selection in Setup** | Dynamic steward checkboxes in meeting setup dialog with Select All / Clear All | `getStewardEmailsForMeetingSetup()` | steward, select, checkboxes |

### Meeting Notes Dashboard Tab (v4.6.0)

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Meeting Notes Tab** | New tab in Member Dashboard showing completed meetings chronologically with search | `getUnifiedDashboardData()` | dashboard, tab, meeting notes |
| **View-Only Links** | Members can view meeting notes via read-only Google Doc links after the meeting | Within Member Dashboard | view, read-only, link |
| **Search Meeting Notes** | Client-side search/filter across meeting names and dates | `filterMeetingNotes()` | search, filter, notes |

### Meeting Check-In Log Columns (A-P)

| Column | Letter | Description |
|--------|--------|-------------|
| **Meeting ID** | A | Unique meeting identifier |
| **Meeting Name** | B | Display name for the meeting |
| **Meeting Date** | C | Scheduled date |
| **Meeting Type** | D | Virtual / In-Person / Hybrid |
| **Member ID** | E | Checked-in member ID |
| **Member Name** | F | Checked-in member name |
| **Check-In Time** | G | Timestamp of check-in |
| **Email** | H | Member email |
| **Meeting Time** | I | Scheduled time of meeting |
| **Meeting Duration** | J | Duration in minutes |
| **Event Status** | K | Calendar event status |
| **Notify Stewards** | L | Steward emails for attendance reports |
| **Calendar Event ID** | M | Google Calendar event ID |
| **Notes Doc URL** | N | Meeting Notes Google Doc URL |
| **Agenda Doc URL** | O | Meeting Agenda Google Doc URL |
| **Agenda Stewards** | P | Steward emails for early agenda sharing (3 days prior) |

---

## 23. Member Drive Folders

> **File:** `05_Integrations.gs`, `03_UIComponents.gs`

Quick Action to create or reuse a Google Drive folder for any member.

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Create Member Folder** | Creates a Google Drive folder for a member, or reuses an existing grievance folder | `setupDriveFolderForMember(memberId)` | folder, drive, member |
| **Quick Action Button** | "Create Member Folder" button in Member Quick Actions dialog | Within `showMemberQuickActions()` | button, quick action, UI |

---

## 24. Constant Contact Integration

> **File:** `05_Integrations.gs` | **Version:** 4.9.0 | **Type:** Read-Only External API

Pulls email engagement metrics from Constant Contact v3 API into the Member Directory. Matches CC contacts to members by email address (case-insensitive) and populates the `OPEN_RATE` and `RECENT_CONTACT_DATE` columns.

### Setup & Authorization

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **API Credential Setup** | Store Constant Contact API key and client secret in Script Properties (encrypted at rest). One-time setup. | Admin > Data Sync > CC Setup: API Credentials | `showConstantContactSetup()`, API key, client secret, setup |
| **OAuth2 Authorization** | Interactive dialog that walks through the CC OAuth2 flow. Opens auth URL, user grants access, pastes redirect URL to complete token exchange. | Admin > Data Sync > CC Authorize Account | `authorizeConstantContact()`, OAuth, token, authorize |
| **Connection Status** | Shows current API key status, access token validity, expiry time, and refresh token availability. | Admin > Data Sync > CC Connection Status | `showConstantContactStatus()`, status, token, expiry |
| **Disconnect** | Removes all stored CC credentials and tokens. Requires confirmation. | Admin > Data Sync > CC Disconnect | `disconnectConstantContact()`, remove, credentials, disconnect |

### Engagement Sync

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Sync CC Engagement** | Fetches all CC contacts, matches by email to Member Directory, pulls per-contact activity summaries (open rate, last activity date), writes to `OPEN_RATE` and `RECENT_CONTACT_DATE` columns. Progress toasts during sync. | Admin > Data Sync > Sync CC Engagement → Members | `syncConstantContactEngagement()`, open rate, email, engagement, sync |
| **Automatic Token Refresh** | Access tokens (2-hour lifetime) are automatically refreshed using the stored refresh token. No user action needed after initial authorization. | Automatic | `getConstantContactToken_()`, refresh, token, auto |
| **Rate Limiting** | Respects CC API limits (4 requests/second, 10,000/day). Pauses between batches of API calls. | Automatic | rate limit, throttle, pause |
| **Pagination** | Handles CC accounts with large contact lists (500+ contacts) via cursor-based pagination. | Automatic | pagination, cursor, large list |

### Data Flow

| Source (Constant Contact) | Calculation | Target (Member Directory) |
|--------------------------|-------------|--------------------------|
| `activity_summary.em_sends` + `em_opens` | `(opens / sends) × 100` | **OPEN_RATE** (column T) — Email open rate % |
| `activity_summary.em_sends_date`, `em_opens_date`, `em_clicks_date` | Most recent date | **RECENT_CONTACT_DATE** (column Y) — Last email activity |

### Technical Details

| Property | Value |
|----------|-------|
| **API Base URL** | `https://api.cc.email/v3` |
| **Auth Method** | OAuth 2.0 Authorization Code Grant |
| **Token Storage** | Script Properties (encrypted at rest by Google) |
| **Lookback Period** | 365 days of campaign activity |
| **Matching Key** | Email address (case-insensitive) |
| **Direction** | Read-only — never writes to Constant Contact |

### Requirements

- Constant Contact paid account (Lite $12/mo, Standard $35/mo, or Premium $80/mo)
- CC API application created at the CC Developer Portal
- API key (client ID) and client secret
- One-time OAuth2 authorization via browser

---

## 25. Workload Tracker

> **Files:** `25_WorkloadService.gs` (SPA module) | **Version:** 4.10.0 / 4.13.0
>
> **Note:** The standalone PIN-auth portal (`18_WorkloadTracker.gs` / `WorkloadTracker.html`) is DDS-Dashboard only. Union-Tracker uses the SPA workload module exclusively.

### SPA Workload Module (v4.13.0)

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **SPA Workload Form** | SSO-authenticated workload submission embedded in member_view | `WorkloadService.getFormData()` | SPA, SSO, form |
| **SPA Workload Submit** | Submit workload via SPA without separate PIN auth | `WorkloadService.submitData()` | submit, SPA, SSO |

### Workload Categories

| Category | Description |
|----------|-------------|
| Priority Cases | High-priority cases requiring immediate attention |
| Pending Cases | Cases awaiting action or response |
| Unread Documents | Documents not yet reviewed |
| To-Do Items | Outstanding tasks |
| Sent Referrals | Referrals sent to other departments |
| CE Activities | Continuing education activities |
| Assistance Requests | Requests for assistance received |
| Aged Cases | Cases open beyond expected timeframes |

### Privacy Controls

| Level | Description |
|-------|-------------|
| Unit Anonymous | Identity hidden within unit; data included in collective stats |
| Agency Anonymous | Identity hidden agency-wide; data included in collective stats |
| Private | Data excluded from all collective reporting |

### Workload Sheets

| Sheet | Purpose |
|-------|---------|
| `_Workload_Vault` | Encrypted raw submission data |
| `_Workload_Reporting` | Anonymized collective statistics |
| `_Workload_Reminders` | Email reminder configuration |
| `_Workload_UserMeta` | Member preferences and metadata |
| `_Workload_Archive` | Archived data (24-month rolling) |

---

## 26. Resources Hub

> **File:** `08a_SheetSetup.gs`, SPA routes | **Version:** 4.11.0

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Resources Page** | Educational content hub with search, category pills, expandable cards | `?page=resources` route | resources, articles, education |
| **Resource List API** | Returns visible resources with audience filtering | `getWebAppResourcesList()` | API, filter, audience |
| **Category Navigation** | Category pill buttons for filtering content | Within SPA | categories, filter, navigation |
| **Starter Articles** | 8 pre-loaded articles: Know Your Rights, Grievance Process, FAQ, Forms & Templates | Via `CREATE_DASHBOARD()` | starter, articles, setup |

### Resources Sheet (12 columns)

| Column | Description |
|--------|-------------|
| Resource_ID | Unique identifier (RES-XXX) |
| Title | Resource title |
| Category | Article category |
| Content | Full article content |
| Audience | Target audience (All, Members, Stewards) |
| Status | Active / Inactive |
| Created_Date | Date created |
| Created_By | Author |
| Updated_Date | Last update date |
| Sort_Order | Display order |
| Tags | Searchable tags |
| Visibility | Visible / Hidden |

---

## 27. Notifications System

> **File:** `08a_SheetSetup.gs`, SPA routes | **Version:** 4.12.0

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Get Notifications** | Filters active, non-expired, non-dismissed, audience-matched notifications | `getWebAppNotifications(email, role)` | notifications, filter, active |
| **Dismiss Notification** | Per-member dismiss tracking via Dismissed_By column | `dismissWebAppNotification(id, email)` | dismiss, hide, per-member |
| **Send Notification** | Steward compose with auto-ID (NOTIF-XXX) | `sendWebAppNotification(data)` | send, compose, steward |
| **Recipient List** | Member directory + preset groups for targeting | `getNotificationRecipientList()` | recipients, groups, targeting |
| **Notifications Page** | Dual-role page: member cards + steward inline compose | `?page=notifications` route | page, cards, compose |

### Notification Types & Priority

| Type | Description |
|------|-------------|
| Steward Message | Direct message from steward to members |
| Announcement | Organization-wide announcement |
| Deadline | Deadline reminder or alert |
| System | System-generated notification |

| Priority | Behavior |
|----------|----------|
| Normal | Standard display order |
| Urgent | Sorts first, highlighted display |

---

## 28. SPA Web Dashboard

> **Files:** `19_WebDashAuth.gs`, `20_WebDashConfigReader.gs`, `21_WebDashDataService.gs`, `22_WebDashApp.gs`, `23_PortalSheets.gs`, `24_WeeklyQuestions.gs` | **Version:** 4.12.2

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **SSO Authentication** | Google SSO with automatic role detection | `initWebDashboardAuth()` | SSO, Google, auth |
| **Magic Link Auth** | Email-based authentication for non-Google users | `verifyMagicLink(token)` | magic link, email, token |
| **Steward View** | Full steward dashboard in SPA format | `steward_view.html` | steward, dashboard, SPA |
| **Member View** | Member-facing dashboard in SPA format | `member_view.html` | member, dashboard, SPA |
| **Deep-Link Routing** | `?page=X` pre-selects tabs via `PAGE_DATA.initialTab` | `doGetWebDashboard(e)` | deep-link, routing, tabs |
| **Config Reader** | Column-based Config tab reader with `CONFIG_COLS` | `getConfigValue(key)` | config, settings, reader |
| **Weekly Questions** | Weekly check-in questions with response tracking | `getWeeklyQuestions()`, `submitWeeklyResponse()` | weekly, questions, check-in |
| **Portal Sheets** | Hidden sheet management for SPA data | `getPortalSheetData()` | sheets, data, hidden |

### SPA Hidden Sheets

| Sheet | Columns | Purpose |
|-------|---------|---------|
| `_Weekly_Questions` | varies | Weekly check-in questions and responses |
| `_Contact_Log` | 8 | Member contact tracking |
| `_Steward_Tasks` | 10 | Steward task management |

---

## 29. Notification Bell & EventBus Alerts

> **Files:** `15_EventBus.gs`, SPA views | **Version:** 4.13.0

| Feature | Description | Function | Keywords |
|---------|-------------|----------|----------|
| **Notification Bell** | Bell icon with unread count badge in SPA header | Within SPA views | bell, badge, unread |
| **Steward Notification Management** | Compose/inbox/manage tabs in steward view | Within `steward_view.html` | compose, inbox, manage |
| **EventBus Auto-Notifications** | Automatic alerts for grievance deadlines and status changes | `EventBus.emit('notification:auto')` | auto, EventBus, alerts |
| **Member Notification View** | View and dismiss notifications in member view | Within `member_view.html` | view, dismiss, member |

---

## Version History

| Version | Features Added |
|---------|----------------|
| **4.13.0** | Notification bell with unread badge, EventBus auto-notifications, WorkloadService SPA module |
| **4.12.2** | SPA web dashboard — SSO + magic link auth, steward/member views, deep-link routing |
| **4.12.0** | Notifications system — sheet, API, dual-role page, steward compose |
| **4.11.0** | Resources hub, meeting check-in route, design refresh (DM Sans + Fraunces) |
| **4.10.0** | Workload Tracker — 8 categories, privacy controls, reciprocity, email reminders |
| **4.9.0** | Constant Contact v3 API integration — read-only email engagement metrics sync (OPEN_RATE, RECENT_CONTACT_DATE) |
| **4.8.2** | State field added to member contact update (self-service portal, contact form, profile) |
| **4.6.0** | Meeting Notes & Agenda doc automation, two-tier steward agenda sharing, Meeting Notes dashboard tab, member Drive folders, meeting event scheduling, grievance date override |
| **4.5.1** | Engagement tracking fixes, 950 Jest tests, GRIEVANCE_OUTCOMES/generateGrievanceId fixes |
| **4.5.0** | Security module, Data Access Layer, Member Self-Service PIN, consolidated architecture |
| **4.4.1** | Dynamic Engine, Grievance Reminders, Looker Studio Integration (Standard & PII-Free) |
| **4.3.8** | Searchable Help Guide, Features Reference Sheet, Enhanced FAQ |
| **4.3.7** | Dynamic row styling with getMaxRows() |
| **4.3.6** | Full row styling for zebra stripes |
| **4.3.5** | Production polish, NUKE enhancements, tab colors |
| **4.3.4** | Member Satisfaction 8-section analysis |
| **4.3.3** | Two-Dashboard Architecture (Steward + Member) |
| **4.2.0** | Modal SPA architecture, Web App entry point |
| **4.0.3** | Material Design, Google Charts, Safety Valve PII |
| **4.0.0** | Strategic Command Center, PDF engine, mobile view |

---

## See Also

- [README.md](README.md) - Quick start and installation
- [AIR.md](AIR.md) - Architecture and implementation reference
- [USER_TUTORIALS.md](USER_TUTORIALS.md) - Step-by-step tutorials
- [INTERACTIVE_DASHBOARD_GUIDE.md](INTERACTIVE_DASHBOARD_GUIDE.md) - Dashboard customization
- [STEWARD_GUIDE.md](STEWARD_GUIDE.md) - Steward-focused features
- [COMFORT_VIEW_GUIDE.md](COMFORT_VIEW_GUIDE.md) - Accessibility options

---

*Strategic Command Center v4.13.0 - A personal project providing comprehensive tools for representatives*
