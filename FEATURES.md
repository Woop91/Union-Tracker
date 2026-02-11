# 509 Strategic Command Center - Complete Features Reference

**Version:** 4.5.1 | **Codename:** Code Audit & Cleanup
**Last Updated:** February 2026

> **New in v4.5.1:** Fixed tab population bugs, added sheet tab colors, 950 Jest tests, engagement tracking fixes, consolidated architecture (16 source files)

This document provides a comprehensive, searchable reference of all features in the 509 Dashboard system. Use `Ctrl+F` (or `Cmd+F` on Mac) to search for specific features.

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

---

## 1. Dashboard & Analytics

### Two-Dashboard Architecture (v4.3.3)

| Feature | Description | Menu Path | Keywords |
|---------|-------------|-----------|----------|
| **Steward Dashboard** | Internal dashboard with 6 tabs: Overview, Workload, Analytics, Hot Spots, Bargaining, Satisfaction. Contains member names and PII - for steward use only. | Strategic Ops > Command Center > Steward Dashboard | internal, analytics, PII, workload |
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
| **Auto-Generate Member IDs** | Creates IDs in format M + First2 + Last2 + 3 digits (e.g., MJOSM123 for John Smith). | Strategic Ops > ID Engines > Generate Missing IDs | ID, generate, auto |
| **Check Duplicate IDs** | Finds and highlights duplicate Member IDs in the directory. | Strategic Ops > ID Engines > Check Duplicates | duplicate, check, validate |
| **Unit Code Configuration** | Custom unit codes for ID generation (e.g., Main Station:MS). | Config sheet column AT | unit, code, config |

### Member Directory Columns

| Column | Description | Auto-Calculated? |
|--------|-------------|------------------|
| **Member ID** | Unique identifier (MJOHN123 format) | Optional auto-gen |
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
| **Create Calendar Event** | Manually create calendar event for selected grievance. | 509 Dashboard > Calendar > Create Event | create, event, manual |

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
| **Apply Global Styling** | Applies Roboto font and zebra stripes to all rows in data sheets. | 509 Dashboard > Styling > Apply Global | styling, font, Roboto |
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
| **Pocket/Mobile View** | Hides non-essential columns for phone access. | 509 Dashboard > Field Access > Mobile View | mobile, pocket, phone |
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
| `_Looker_Satisfaction` | Member Satisfaction | Survey responses with section averages and scores |

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
| **PIN Entry Dialog** | Secure PIN entry interface for member authentication | `showPINEntryDialog()` | dialog, PIN, entry |

---

## Version History

| Version | Features Added |
|---------|----------------|
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

*509 Strategic Command Center v4.5.0 - A personal project providing comprehensive tools for representatives*
