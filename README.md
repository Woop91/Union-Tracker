# 509 Strategic Command Center

**Version 4.7.0** | Union Steward Dashboard for Google Sheets

A Google Sheets-based system for managing union grievances, tracking member records, monitoring deadlines, and running steward operations. Built on Google Apps Script with a 27-file modular architecture.

---

## Table of Contents

- [What It Does](#what-it-does)
- [Quick Start](#quick-start)
- [Installation](#installation)
- [Going Live (Production)](#going-live-production)
- [Key Features](#key-features)
- [Menu System](#menu-system)
- [Deadline Rules (Article 23A)](#deadline-rules-article-23a)
- [Sheets Structure](#sheets-structure)
- [Development](#development)
- [Architecture](#architecture)
- [Troubleshooting](#troubleshooting)
- [Documentation Index](#documentation-index)
- [Contributing](#contributing)
- [Version History](#version-history)
- [License](#license)

---

## What It Does

The 509 Dashboard gives union stewards and leadership a centralized place to:

- **Track grievances** from filing through resolution with automatic deadline calculations
- **Manage member records** with contact info, job metadata, and steward assignments
- **Monitor deadlines** with color-coded urgency and Google Calendar sync
- **Run dashboards** for executive analytics, member-facing stats, and steward workload
- **Automate meetings** with auto-generated Google Docs for notes and agendas
- **Generate PDFs** for grievance forms with digital signature blocks
- **Send notifications** for overdue cases, escalation alerts, and reminders
- **Protect privacy** with PII masking on all public-facing dashboards

The system runs entirely inside Google Sheets with no external servers required.

---

## Quick Start

If you want to try it out quickly with demo data:

1. Deploy the scripts to Google Apps Script (see [Installation](#installation))
2. Refresh your Google Sheet -- 4 custom menus will appear
3. Go to **Admin > Demo Data > Seed All Sample Data**
4. Explore the dashboard using the **Union Hub** and **Strategic Ops** menus

When you're done testing, run **Admin > Demo Data > NUKE SEEDED DATA** to remove all test data. See the [Seed & Nuke Guide](SEED_NUKE_GUIDE.md) for details.

---

## Installation

### Prerequisites

- Google account with Google Sheets access
- Node.js 18+ and npm (for building and testing locally)
- [clasp](https://github.com/google/clasp) CLI (optional, for automated deployment)

### Option A: Manual Deployment

1. **Clone the repository:**

   ```bash
   git clone https://github.com/Woop91/MULTIPLE-SCRIPS-REPO.git
   cd MULTIPLE-SCRIPS-REPO
   ```

2. **Create a new Google Sheet** at [sheets.google.com](https://sheets.google.com)

3. **Open the Apps Script editor:** Extensions > Apps Script

4. **Copy each file** from the `src/` folder into the Apps Script editor:
   - Click the **+** button next to "Files" to create a new script file
   - Name it to match the source file (without the `.gs` extension)
   - Paste the contents from the corresponding `src/` file
   - Repeat for all 27 `.gs` files and the `.html` file

5. **Run the initial setup:**
   - Select `CREATE_509_DASHBOARD` from the function dropdown
   - Click **Run**
   - Authorize the script when prompted (Advanced > Go to project name > Allow)

6. **Refresh your Google Sheet** -- the menus will appear

### Option B: Automated Deployment with clasp

1. **Clone and install:**

   ```bash
   git clone https://github.com/Woop91/MULTIPLE-SCRIPS-REPO.git
   cd MULTIPLE-SCRIPS-REPO
   npm install
   ```

2. **Login and create the project:**

   ```bash
   clasp login
   clasp create --type sheets --title "Union Dashboard"
   ```

3. **Build and deploy:**

   ```bash
   npm run build:prod
   clasp push
   ```

4. **Open the sheet and refresh** -- run `CREATE_509_DASHBOARD` from the Apps Script editor to initialize.

### First-Time Configuration

After installation, configure the system for your organization:

1. Open the **Config** sheet tab
2. Fill in your organization's data:
   - Column A: Job Titles
   - Column B: Office Locations
   - Column C: Units
   - Column F: Supervisors
   - Column G: Managers
   - Column H: Stewards
3. Run **Admin > System Diagnostics** to verify everything is set up correctly

See [USER_TUTORIALS.md](USER_TUTORIALS.md) for step-by-step walkthroughs.

---

## Going Live (Production)

When you're ready to move from testing to real data:

1. Run **Admin > Demo Data > NUKE SEEDED DATA** to remove all test data
2. Delete `07_DevTools.gs` from your Apps Script project (Extensions > Apps Script > right-click > Delete)
3. Refresh the sheet -- the Demo menu disappears automatically

After that, you have **26 production files** and a clean system ready for real member data. See the [Seed & Nuke Guide](SEED_NUKE_GUIDE.md) for the full process.

---

## Key Features

### Grievance Management
- Full lifecycle tracking: filing, step advancement, resolution
- Auto-calculated deadlines based on Article 23A rules
- Step escalation: Step I > Step II > Step III > Arbitration
- Status tracking with color-coded rows (Open, Pending, Won, Denied, Settled, In Arbitration, Closed)
- Reminder system with two reminder slots per grievance

### Member Management
- Centralized member directory with contact info and job metadata
- Auto-generated Member IDs (format: UNIT_CODE-SEQUENCE-H, e.g., MS-101-H)
- Steward assignments and workload tracking
- Import/export capabilities
- Self-service portal with PIN authentication

### Meeting Management (v4.6.0)
- Auto-created Google Docs for Meeting Notes and Meeting Agenda
- Two-tier steward agenda sharing (selected stewards get it 3 days early, all stewards 1 day before)
- Meeting check-in system with attendance logging
- Meeting Notes dashboard tab with chronological view and search
- Google Calendar integration for scheduling

### Dashboards & Analytics
- **Steward Dashboard**: Internal view with 11 tabs -- Overview, My Cases, Workload, Analytics, Directory, Hot Spots, Bargaining, Satisfaction, Resources, Compare, Meeting Notes
- **Member Dashboard**: PII-safe view for sharing with members
- **Executive Dashboard**: High-level metrics with Chart.js visualizations
- **Interactive Dashboard**: Customizable metrics and chart types
- Hot spot detection, sentiment analysis, and workload balancing

### Drive & Calendar Integration
- Auto-created Drive folders for each grievance (with step subfolders)
- Per-member Drive folders via Quick Actions
- Deadline sync to Google Calendar
- Automated meeting event creation

### Notifications & Alerts
- Daily overdue alerts (8 AM trigger)
- Escalation notifications for Step II/III/Arbitration
- Scheduled meeting document emails
- Email snapshots and dashboard links

### Accessibility (Comfort View)
- ADHD-friendly themes with soft colors
- Focus mode and zebra stripes
- Adjustable font sizes and high-contrast options
- Reduced motion setting
- Gridline toggle

### Security
- XSS prevention with HTML escaping
- PII masking for public dashboards (phone numbers, SSNs auto-redacted)
- Zero-knowledge survey vault: email and member ID stored as SHA-256 hashes only (non-reversible)
- Survey responses are cryptographically anonymous — no one can link answers to members
- Formula injection protection
- Input sanitization and validation
- Audit logging of all changes
- Sabotage detection (mass deletion alerts)

### Looker Studio Integration
- **Standard**: Hidden `_Looker_*` sheets with full data for internal reports
- **PII-Free**: Anonymized `_Looker_Anon_*` sheets for external stakeholders
- Non-reversible hashes, bucketed values, and engagement levels
- Survey data uses zero-knowledge vault — `_Looker_Satisfaction` contains no member IDs or emails

---

## Menu System

The dashboard provides 4 top-level menus:

### Union Hub (Main Menu)

| Submenu | What It Does |
|---------|--------------|
| Search | Desktop, Quick, and Advanced search |
| Grievances | Create, edit, bulk update, view active grievances |
| Members | Add, find, import/export, steward directory, anonymous surveys |
| Calendar | Sync deadlines, view upcoming, clear events |
| Google Drive | Setup folders, view files, batch create |
| Notifications | Email settings, test notifications |
| View | Dashboards, dark mode, themes |
| Comfort View | Focus mode, zebra stripes, font sizes |
| Multi-Select | Multi-select editor, auto-open triggers |

### Admin (System Administration)

| Submenu | What It Does |
|---------|--------------|
| Automation | Auto-refresh, midnight triggers, email snapshots |
| Data Sync | Sync all data, install triggers |
| Validation | Bulk validation, settings, indicators |
| Cache | Cache status, warm up, clear caches |
| Setup | Hidden sheets, data validations, defaults |
| Demo Data | Seed sample data, nuke seeded data |

### Strategic Ops (Strategic Operations)

| Submenu | What It Does |
|---------|--------------|
| Command Center | Steward/Member dashboards, steward performance |
| Strategic Intelligence | Hot zones, rising stars, hostility report |
| Analytics & Charts | Treemap, sentiment trends, workload report |
| ID & Data Engines | ID generation, duplicate check, PDF creation |
| Steward Management | Promote/demote, contact forms, surveys |

### Field Portal (Mobile/Field Operations)

| Submenu | What It Does |
|---------|--------------|
| Field Accessibility | Mobile view, get mobile URL |
| Analytics & Insights | Unit health, grievance trends, precedents |
| Web App & Portal | Build portals, send emails, JSON APIs |

---

## Deadline Rules (Article 23A)

Deadlines are calculated in business days (excluding weekends):

| Step | Action | Days |
|------|--------|------|
| Filing | From incident date | 21 days |
| Step I | Management response | 7 days |
| Step II | Appeal deadline | 7 days |
| Step II | Management response | 14 days |
| Step III | Appeal deadline | 10 days |
| Step III | Management response | 21 days |
| Arbitration | Demand deadline | 30 days |

### Status Colors

| Status | Color | Meaning |
|--------|-------|---------|
| Open | Yellow | Active case |
| Pending Info | Purple | Waiting on information |
| Won | Green | Victory |
| Denied | Red | Loss |
| Settled | Blue | Negotiated resolution |
| In Arbitration | Red | High stakes |
| Closed | Gray | Complete |

---

## Sheets Structure

### Visible Sheets

| Sheet | Purpose |
|-------|---------|
| Config | Dropdown values and organization settings |
| Member Directory | Union member records (34 columns) |
| Grievance Log | Grievance tracking (34+ columns) |
| Dashboard | Summary statistics and metrics |
| Member Satisfaction | Survey response tracking |
| Getting Started | Setup instructions |
| FAQ | Common questions and answers |
| Config Guide | Configuration help |

### Hidden Sheets (Auto-Managed)

These sheets power the auto-updating columns. You don't need to edit them.

| Sheet | Purpose |
|-------|---------|
| `_Dashboard_Calc` | Dashboard summary metrics |
| `_Grievance_Calc` | Grievance data for Member Directory |
| `_Grievance_Formulas` | Self-healing formulas for timeline columns |
| `_Member_Lookup` | Member data for Grievance Log |
| `_Steward_Contact_Calc` | Steward contact calculations |
| `_Steward_Performance_Calc` | Steward performance scores |
| `_Audit_Log` | System activity log |
| `_Checklist_Calc` | Checklist calculations |

---

## Development

### Setup

```bash
git clone https://github.com/Woop91/MULTIPLE-SCRIPS-REPO.git
cd MULTIPLE-SCRIPS-REPO
npm install
```

### Available Commands

```bash
npm run build          # Build consolidated file (dist/ConsolidatedDashboard.gs)
npm run build --prod   # Production build (excludes DevTools)
npm run lint           # ESLint code quality checks
npm run lint:fix       # Auto-fix ESLint issues
npm run test:unit      # Run 950+ Jest unit tests
npm test               # Full pipeline: lint + build + test
npm run clean          # Clean dist directory
npm run deploy         # Deploy to Google Apps Script (requires clasp)
```

### Making Changes

1. Edit files in the `src/` directory
2. Run `npm run lint` to check for issues
3. Run `npm run build` to generate the consolidated file
4. Run `npm run test:unit` to execute tests
5. Copy updated files to Google Apps Script (or use `clasp push`)

See the [Developer Guide](DEVELOPER_GUIDE.md) for architecture details, code patterns, and debugging tips.

---

## Architecture

The codebase uses a 27-file modular architecture with numbered prefixes that indicate load order and purpose:

| Prefix | Layer | Files |
|--------|-------|-------|
| 00 | Foundation | `00_Security.gs`, `00_DataAccess.gs` |
| 01 | Core | `01_Core.gs` (constants, error handling) |
| 02 | Data | `02_DataManagers.gs` |
| 03-04 | UI | `03_UIComponents.gs`, `04a_UIMenus.gs`, `04b_AccessibilityFeatures.gs`, `04c_InteractiveDashboard.gs`, `04d_ExecutiveDashboard.gs`, `04e_PublicDashboard.gs` |
| 05 | Integrations | `05_Integrations.gs` (Drive, Calendar, WebApp) |
| 06 | Maintenance | `06_Maintenance.gs` (diagnostics, cache, undo) |
| 07 | Dev Tools | `07_DevTools.gs` (remove before production) |
| 08 | Sheet Utilities | `08a_SheetSetup.gs`, `08b_SearchAndCharts.gs`, `08c_FormsAndNotifications.gs`, `08d_AuditAndFormulas.gs` |
| 09 | Dashboards | `09_Dashboards.gs` |
| 10 | Business Logic | `10_Main.gs`, `10a_SheetCreation.gs`, `10b_SurveyDocSheets.gs`, `10c_FormHandlers.gs`, `10d_SyncAndMaintenance.gs` |
| 11 | Command Hub | `11_CommandHub.gs` |
| 12 | Features | `12_Features.gs` (Dynamic Engine, Looker) |
| 13 | Self-Service | `13_MemberSelfService.gs` (PIN auth) |
| 14 | Meetings | `14_MeetingCheckIn.gs` |
| -- | HTML | `MultiSelectDialog.html` |

### Design Principles

- **Separation of Concerns**: Each file has one clear purpose
- **Numbered Prefixes**: Show dependency order for build concatenation
- **Production Ready**: Delete `07_DevTools.gs` for a 26-file production deployment
- **Failure Isolation**: A bug in Calendar sync won't break the Member Directory
- **Self-Healing**: Hidden calculation sheets auto-repair their formulas
- **Performance**: CacheService integration and batch operations handle 5,000+ members

---

## Troubleshooting

### Running Diagnostics

1. Go to **Admin > System Diagnostics**
2. Review the diagnostic report for errors and warnings
3. Click **Run Repair** to fix common issues

### Common Issues

| Problem | Solution |
|---------|----------|
| Menus don't appear | Refresh the page (F5), or run `onOpen()` from the Apps Script editor |
| Missing sheets | Run **Admin > Setup > Hidden Sheets** or `CREATE_509_DASHBOARD()` |
| Broken formulas | Run **Admin > Repair Dashboard** |
| Calendar not syncing | Check Calendar permissions in Settings |
| Columns not auto-updating | Run **Admin > Data Sync > Sync All Data** |
| Dashboard shows stale data | Run **Admin > Cache > Clear All Caches** |
| Slow performance | Clear caches and reduce data in Member Directory |
| Authorization error | Re-authorize: run any function from the Apps Script editor and follow the prompts |
| Hidden sheets missing | Run **Admin > Setup > Setup All Hidden Sheets** |
| Demo menu won't go away | Close and reopen the spreadsheet after running NUKE |

### FAQ

**Q: How many members can it handle?**
A: The system is optimized for 5,000+ members using batch operations and caching.

**Q: Can I use this on mobile?**
A: Yes. Use **Field Portal > Field Accessibility > Mobile View** for a phone-optimized layout, or access the web app portal.

**Q: Is member data visible to everyone?**
A: No. The Member Dashboard uses PII masking to hide personal information. Only the Steward Dashboard shows full member data.

**Q: Can I undo changes?**
A: The system has a 50-action undo/redo history. You can also check the audit log for change tracking.

**Q: How do I back up my data?**
A: Use Google Sheets' built-in version history (File > Version history) or make a copy of the spreadsheet (File > Make a copy).

---

## Documentation Index

| Document | Description |
|----------|-------------|
| [FEATURES.md](FEATURES.md) | Complete searchable feature reference (23 categories) |
| [USER_TUTORIALS.md](USER_TUTORIALS.md) | Step-by-step tutorials for common tasks |
| [STEWARD_GUIDE.md](STEWARD_GUIDE.md) | Guide for stewards using the system |
| [DEVELOPER_GUIDE.md](DEVELOPER_GUIDE.md) | Architecture, code patterns, and debugging |
| [CONTRIBUTING.md](CONTRIBUTING.md) | How to contribute to the project |
| [CHANGELOG.md](CHANGELOG.md) | Detailed version history |
| [INTERACTIVE_DASHBOARD_GUIDE.md](INTERACTIVE_DASHBOARD_GUIDE.md) | Dashboard customization guide |
| [COMFORT_VIEW_GUIDE.md](COMFORT_VIEW_GUIDE.md) | Accessibility and visual comfort features |
| [GRIEVANCE_WORKFLOW_GUIDE.md](GRIEVANCE_WORKFLOW_GUIDE.md) | Grievance filing workflow |
| [SEED_NUKE_GUIDE.md](SEED_NUKE_GUIDE.md) | Demo data seeding and cleanup |
| [QUICK_DEPLOY.md](QUICK_DEPLOY.md) | Fast deployment instructions |
| [AIR.md](AIR.md) | Full architecture and implementation reference |
| [SECURITY_REVIEW.md](SECURITY_REVIEW.md) | Security analysis and findings |

### Setup Guides

| Guide | Description |
|-------|-------------|
| [Heatmap Setup](setup-instructions/01_HEATMAP_SETUP.md) | Color gradient configuration |
| [OCR Setup](setup-instructions/02_OCR_SETUP.md) | Google Cloud Vision API for OCR |
| [CLASP Deployment](setup-instructions/03_CLASP_SETUP.md) | Complete CLASP deployment guide |
| [Drive URL Setup](setup-instructions/04_RESOURCE_DRIVE_URL_LINK_SETUP.md) | Shared documents folder linking |
| [Dropbox Setup](setup-instructions/05_RESOURCE_DROPBOX_SETUP.md) | Resource dropbox folder setup |
| [PDF Generation](setup-instructions/06_PDF_GENERATION_SETUP.md) | PDF template and email setup |

---

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/your-feature`)
3. Edit files in `src/`
4. Run `npm test` (lint + build + tests)
5. Commit with descriptive messages
6. Submit a pull request

See [CONTRIBUTING.md](CONTRIBUTING.md) for full guidelines.

---

## Version History

| Version | Date | Highlights |
|---------|------|------------|
| **4.7.0** | 2026-02-14 | Security hardening, 40+ code review fixes, 1090 tests across 20 suites, deduplicated escapeHtml, lint clean |
| **4.6.0** | 2026-02-12 | Meeting Notes & Agenda doc automation, two-tier steward agenda sharing, Meeting Notes dashboard tab, member Drive folders, meeting event scheduling |
| **4.5.1** | 2026-02-11 | Engagement tracking fixes, 950 Jest unit tests, critical bug fixes |
| **4.5.0** | 2026-02-01 | Security module, Data Access Layer, Member Self-Service with PIN auth, consolidated 27-file architecture |
| **4.4.1** | 2026-01-31 | Dynamic Engine, Looker Studio integration, Grievance Reminders |
| **4.4.0** | 2026-01-30 | Grievance tracking, dashboards, satisfaction surveys, calendar integration |
| **4.3.x** | 2026-01 | Searchable Help Guide, two-dashboard architecture, production polish |
| **4.2.0** | 2026-01 | Modal SPA architecture, Web App entry point, member portal |
| **4.0.x** | 2026-01 | Strategic Command Center, PDF engine, mobile view, Material Design |

See [CHANGELOG.md](CHANGELOG.md) for detailed release notes.

---

## License

MIT License - Free for personal use. See LICENSE file for details.

**Creator:** Wardis N. Vizcaino (wardis@pm.me)
